"""严格校验排版后的 docx 是否满足用户规范。"""
import sys, re, zipfile
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def W(tag): return '{%s}%s' % (NS['w'], tag)
def R(tag): return '{%s}%s' % (NS['r'], tag)

def load_xml(zf, name):
    with zf.open(name) as f:
        return etree.parse(f)

def get_run_props(run, styles_xml, para_style_id):
    effective = {'font': None, 'size': None, 'bold': None}
    rpr = run.find(W('rPr'))
    if rpr is not None:
        fonts = rpr.find(W('rFonts'))
        if fonts is not None:
            ea = fonts.get(W('eastAsia'))
            if ea: effective['font'] = ea
            elif fonts.get(W('ascii')): effective['font'] = fonts.get(W('ascii'))
        sz = rpr.find(W('sz'))
        if sz is not None and sz.get(W('val')):
            effective['size'] = sz.get(W('val'))
        b = rpr.find(W('b'))
        if b is not None:
            val = b.get(W('val'))
            effective['bold'] = (val is None or val.lower() in ('true', '1'))
    if para_style_id and any(v is None for v in effective.values()):
        style = find_style(styles_xml, para_style_id)
        if style is not None:
            s_rpr = style.find(W('rPr'))
            if s_rpr is not None:
                fonts = s_rpr.find(W('rFonts'))
                if fonts is not None and effective['font'] is None:
                    ea = fonts.get(W('eastAsia')) or fonts.get(W('ascii'))
                    if ea: effective['font'] = ea
                sz = s_rpr.find(W('sz'))
                if sz is not None and effective['size'] is None and sz.get(W('val')):
                    effective['size'] = sz.get(W('val'))
                b = s_rpr.find(W('b'))
                if b is not None and effective['bold'] is None:
                    val = b.get(W('val'))
                    effective['bold'] = (val is None or val.lower() in ('true', '1'))
    return effective

def find_style(styles_xml, style_id):
    root = styles_xml.getroot()
    for st in root.findall(W('style')):
        if st.get(W('styleId')) == style_id:
            return st
    return None

def get_para_props(p, styles_xml):
    result = {'line': None, 'lineRule': None, 'firstLineChars': None,
              'justify': None, 'before': None, 'after': None, 'styleId': None}
    ppr = p.find(W('pPr'))
    if ppr is not None:
        pstyle = ppr.find(W('pStyle'))
        if pstyle is not None: result['styleId'] = pstyle.get(W('val'))
        spacing = ppr.find(W('spacing'))
        if spacing is not None:
            result['line'] = spacing.get(W('line'))
            result['lineRule'] = spacing.get(W('lineRule'))
            result['before'] = spacing.get(W('before'))
            result['after'] = spacing.get(W('after'))
        ind = ppr.find(W('ind'))
        if ind is not None:
            result['firstLineChars'] = ind.get(W('firstLineChars'))
        jc = ppr.find(W('jc'))
        if jc is not None:
            result['justify'] = jc.get(W('val'))
    if result['styleId']:
        style = find_style(styles_xml, result['styleId'])
        if style is not None:
            s_ppr = style.find(W('pPr'))
            if s_ppr is not None:
                spacing = s_ppr.find(W('spacing'))
                if spacing is not None:
                    if result['line'] is None: result['line'] = spacing.get(W('line'))
                    if result['lineRule'] is None: result['lineRule'] = spacing.get(W('lineRule'))
                    if result['before'] is None: result['before'] = spacing.get(W('before'))
                    if result['after'] is None: result['after'] = spacing.get(W('after'))
                ind = s_ppr.find(W('ind'))
                if ind is not None and result['firstLineChars'] is None:
                    result['firstLineChars'] = ind.get(W('firstLineChars'))
                jc = s_ppr.find(W('jc'))
                if jc is not None and result['justify'] is None:
                    result['justify'] = jc.get(W('val'))
    return result


SPECS = {
    'Title': {'font': '方正小标宋简体', 'size': '44', 'bold': False, 'justify': 'center'},
    'H1':    {'font': '黑体', 'size': '32', 'bold': False,
              'line': '360', 'firstLineChars': '200', 'before': '120', 'after': '120'},
    'H2':    {'font': '楷体_GB2312', 'size': '32', 'bold': True,
              'line': '360', 'firstLineChars': '200'},
    'H3':    {'font': '仿宋_GB2312', 'size': '32', 'bold': True,
              'line': '360', 'firstLineChars': '200'},
    'Body':  {'font': '仿宋_GB2312', 'size': '32', 'bold': False,
              'line': '360', 'firstLineChars': '200'},
    'Footnote': {'font': '仿宋_GB2312', 'size': '28', 'bold': False,
              'line': '240'},
}

STYLEID_TO_ROLE = {
    'Title': 'Title',           # 用 Word 内置 styleId，导航窗格正确识别
    'Heading1': 'H1',
    'Heading2': 'H2',
    'Heading3': 'H3',
    'GongWenBody': 'Body',
    'GongWenFootnote': 'Footnote',
}


def footer_analyze(zf, fx):
    """返回 footer 的类型: 'empty' | 'arabic' | 'roman' | 'unknown'。"""
    root = fx.getroot()
    instrs = [r.text or '' for r in root.findall('.//' + W('instrText'))]
    joined = ''.join(instrs)
    if 'PAGE' not in joined:
        return 'empty'
    if 'ROMAN' in joined.upper():
        return 'roman'
    return 'arabic'


def verify(docx_path, expect_cover=False, expect_toc_section=False, min_roles=None):
    failures = []

    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = load_xml(zf, 'word/document.xml')
        styles_xml = load_xml(zf, 'word/styles.xml')

        # ===== Footer 类型映射 =====
        footer_files = sorted([n for n in zf.namelist() if n.startswith('word/footer') and n.endswith('.xml')])
        footer_types = {}
        for fp in footer_files:
            fx = load_xml(zf, fp)
            footer_types[fp] = footer_analyze(zf, fx)

        # Footer 关系映射：rel Id -> file name
        rels = load_xml(zf, 'word/_rels/document.xml.rels')
        rel_target = {}
        for rel in rels.getroot():
            rid = rel.get('Id')
            target = rel.get('Target') or ''
            target = target.lstrip('/')
            if not target.startswith('word/'):
                target = 'word/' + target
            rel_target[rid] = target

        # ===== 段落逐个检查 =====
        body = doc_xml.getroot().find(W('body'))
        paragraphs = body.findall(W('p'))
        role_counts = {'Title': 0, 'H1': 0, 'H2': 0, 'H3': 0, 'Body': 0, 'Footnote': 0}
        body_samples = {'Title': [], 'H1': [], 'H2': [], 'H3': [], 'Body': [], 'Footnote': []}
        untouched_count = 0  # 未被我们的样式覆盖的段落（目录项、封面原样等）

        for pi, p in enumerate(paragraphs):
            pp = get_para_props(p, styles_xml)
            sid = pp['styleId']
            text = ''.join([t.text or '' for t in p.iter(W('t'))])
            if sid not in STYLEID_TO_ROLE:
                untouched_count += 1
                continue
            role = STYLEID_TO_ROLE[sid]
            role_counts[role] = role_counts.get(role, 0) + 1
            if len(body_samples[role]) < 5:
                body_samples[role].append((pi, text[:40]))
            spec = SPECS[role]

            if 'line' in spec and pp['line'] != spec['line']:
                failures.append(f"[P{pi} {role}] line={pp['line']} 期望 {spec['line']} | 文本: {text[:30]}")
            if 'firstLineChars' in spec and pp['firstLineChars'] != spec['firstLineChars']:
                failures.append(f"[P{pi} {role}] firstLineChars={pp['firstLineChars']} 期望 {spec['firstLineChars']} | 文本: {text[:30]}")
            if 'justify' in spec and pp['justify'] != spec['justify']:
                failures.append(f"[P{pi} {role}] justify={pp['justify']} 期望 {spec['justify']} | 文本: {text[:30]}")
            if 'before' in spec and pp['before'] != spec['before']:
                failures.append(f"[P{pi} {role}] before={pp['before']} 期望 {spec['before']} | 文本: {text[:30]}")
            if 'after' in spec and pp['after'] != spec['after']:
                failures.append(f"[P{pi} {role}] after={pp['after']} 期望 {spec['after']} | 文本: {text[:30]}")

            # 段落级回归检查：不能残留 numPr / 主题字体
            ppr_el = p.find(W('pPr'))
            if ppr_el is not None:
                if ppr_el.find(W('numPr')) is not None:
                    failures.append(f"[P{pi} {role}] 残留了 numPr 编号列表 | 文本: {text[:30]}")
                # 段落标记 rPr 里的主题字体也不应该留下
                mark_rpr = ppr_el.find(W('rPr'))
                if mark_rpr is not None:
                    mark_fonts = mark_rpr.find(W('rFonts'))
                    if mark_fonts is not None:
                        for attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                            if mark_fonts.get(W(attr)) is not None:
                                failures.append(f"[P{pi} {role}] 段落标记残留 {attr} 主题字体 | 文本: {text[:30]}")
                                break

            for ri, run in enumerate(p.findall(W('r'))):
                eff = get_run_props(run, styles_xml, sid)
                if not eff['font'] and not eff['size']:
                    continue
                # 回归检查：run 的 rFonts 不能有 *Theme
                rpr = run.find(W('rPr'))
                if rpr is not None:
                    fonts = rpr.find(W('rFonts'))
                    if fonts is not None:
                        for attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                            if fonts.get(W(attr)) is not None:
                                failures.append(f"[P{pi} R{ri} {role}] run 残留 {attr} 主题字体 | 文本: {text[:30]}")
                                break
                if eff['font'] != spec['font']:
                    failures.append(f"[P{pi} R{ri} {role}] font={eff['font']} 期望 {spec['font']} | 文本: {text[:30]}")
                if eff['size'] != spec['size']:
                    failures.append(f"[P{pi} R{ri} {role}] size={eff['size']} 期望 {spec['size']} | 文本: {text[:30]}")
                if eff['bold'] != spec['bold']:
                    failures.append(f"[P{pi} R{ri} {role}] bold={eff['bold']} 期望 {spec['bold']} | 文本: {text[:30]}")

        # ===== SectPr 分析 =====
        all_sectPrs = doc_xml.getroot().findall('.//' + W('sectPr'))
        section_info = []
        for sp in all_sectPrs:
            fref = sp.find(W('footerReference'))
            footer_file = None
            footer_type = 'no-ref'
            if fref is not None:
                rid = fref.get(R('id'))
                if rid in rel_target:
                    footer_file = rel_target[rid]
                    footer_type = footer_types.get(footer_file, 'unknown')
            pgType = sp.find(W('pgNumType'))
            pg_fmt = pgType.get(W('fmt')) if pgType is not None else None
            pg_start = pgType.get(W('start')) if pgType is not None else None
            section_info.append({
                'footer_type': footer_type,
                'pg_fmt': pg_fmt,
                'pg_start': pg_start,
            })

        # 必须至少有一个 arabic-decimal-start1 section（正文）
        arabic_sects = [s for s in section_info if s['footer_type'] == 'arabic'
                        and s['pg_fmt'] == 'decimal' and s['pg_start'] == '1']
        if not arabic_sects:
            failures.append(f"缺少正文 section（arabic footer + decimal + start=1）。当前：{section_info}")

        if expect_cover:
            cover_sects = [s for s in section_info if s['footer_type'] == 'empty']
            if not cover_sects:
                failures.append(f"期望有封面/目录 section（empty footer），但没找到。当前：{section_info}")

        if expect_toc_section:
            roman_sects = [s for s in section_info if s['footer_type'] == 'roman']
            if not roman_sects:
                failures.append(f"期望有目录 section（roman footer），但没找到。当前：{section_info}")

        # ===== 页码 footer 字体/字号 =====
        for fp, ftype in footer_types.items():
            if ftype == 'empty':
                continue
            fx = load_xml(zf, fp)
            for r in fx.getroot().findall('.//' + W('r')):
                instr = r.find(W('instrText'))
                if instr is None or 'PAGE' not in (instr.text or ''):
                    continue
                rpr = r.find(W('rPr'))
                ea = None
                szv = None
                if rpr is not None:
                    fonts = rpr.find(W('rFonts'))
                    if fonts is not None: ea = fonts.get(W('eastAsia'))
                    sz = rpr.find(W('sz'))
                    if sz is not None: szv = sz.get(W('val'))
                if ea != '仿宋_GB2312':
                    failures.append(f"[{fp}] 页码字体={ea} 期望 仿宋_GB2312")
                if szv != '24':
                    failures.append(f"[{fp}] 页码字号={szv} 期望 24 (小四)")
            jcs = fx.getroot().findall('.//' + W('jc'))
            if not any(j.get(W('val')) == 'center' for j in jcs):
                failures.append(f"[{fp}] 页码未居中")

        # ===== Sanity 角色计数 =====
        if min_roles:
            for r, n in min_roles.items():
                if role_counts.get(r, 0) < n:
                    failures.append(f"Sanity: {r} 出现次数 {role_counts.get(r,0)} < 期望最小 {n}")

        # ===== 汇总 =====
        print("="*60)
        print(f"校验文档：{docx_path}")
        print("="*60)
        print(f"段落分类统计（套了我方样式的）:")
        for r in ['Title', 'H1', 'H2', 'H3', 'Body', 'Footnote']:
            c = role_counts[r]
            print(f"  {r:10s} × {c}")
            for pi, t in body_samples[r][:3]:
                print(f"      · P{pi}: {t}")
        print(f"未套我方样式的段落数（封面 / 目录项 / 空段）: {untouched_count}")
        print(f"\nSection 总数: {len(section_info)}")
        for i, s in enumerate(section_info):
            print(f"  #{i}  footer={s['footer_type']:7s}  fmt={s['pg_fmt']}  start={s['pg_start']}")
        print(f"\nFooter 文件类型:")
        for fp, t in footer_types.items():
            print(f"  {fp}  → {t}")
        print()
        if failures:
            print(f"❌ 失败 {len(failures)} 条：")
            for f in failures[:40]:
                print(f"  {f}")
            if len(failures) > 40:
                print(f"  ...（剩余 {len(failures)-40} 条略）")
            return False
        else:
            print("✅ 全部通过")
            return True


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('path')
    ap.add_argument('--expect-cover', action='store_true')
    ap.add_argument('--no-expect-cover', action='store_true')
    ap.add_argument('--expect-toc-section', action='store_true')
    ap.add_argument('--min-title', type=int, default=1)
    ap.add_argument('--min-h1', type=int, default=0)
    ap.add_argument('--min-h2', type=int, default=0)
    ap.add_argument('--min-h3', type=int, default=0)
    ap.add_argument('--min-body', type=int, default=10)
    args = ap.parse_args()

    expect_cover = args.expect_cover and not args.no_expect_cover
    min_roles = {'Title': args.min_title, 'H1': args.min_h1,
                 'H2': args.min_h2, 'H3': args.min_h3, 'Body': args.min_body}
    ok = verify(args.path,
                expect_cover=expect_cover,
                expect_toc_section=args.expect_toc_section,
                min_roles=min_roles)
    sys.exit(0 if ok else 1)
