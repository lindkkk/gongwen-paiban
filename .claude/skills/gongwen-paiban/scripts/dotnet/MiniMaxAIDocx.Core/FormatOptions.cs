namespace MiniMaxAIDocx.Core;

/// <summary>
/// 排版选项。核心行为默认不依赖这些——全部留空就是自动检测。
/// 三个 Marker 参数接受用户随手输入的标题编号样例（"一、" / "1." / "（1）" 等），
/// 由 <see cref="MarkerPatternInferrer"/> 转成宽松 regex，在分类时优先使用。
/// </summary>
public sealed class FormatOptions
{
    /// <summary>原稿来源："wps" / "office" / "auto"（默认）。目前仅做信息标记，不改变清洗力度。</summary>
    public string Source { get; set; } = "auto";

    public string? H1MarkerSample { get; set; }
    public string? H2MarkerSample { get; set; }
    public string? H3MarkerSample { get; set; }
}
