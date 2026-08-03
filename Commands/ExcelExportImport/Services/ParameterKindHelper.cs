using System.Collections.Generic;
using Autodesk.Revit.DB;
using Tools28.Commands.ExcelExportImport.Models;
using Tools28.Localization;

namespace Tools28.Commands.ExcelExportImport.Services
{
    /// <summary>
    /// パラメータ種別（組み込み/共有/プロジェクト）の判定・ラベル化と、
    /// 同名パラメータを区別するための「曖昧性解消接尾辞」の生成/解析を一元管理する。
    ///
    /// エリアの「用途」のように、1カテゴリに同名パラメータが複数存在する場合、
    /// 名前だけでは区別できないため DisplayName／Excel ヘッダーに接尾辞
    /// （例: <c>I-用途【組み込み】</c> / <c>I-用途【共有】</c>）を付けて区別する。
    /// この接尾辞は Excel を経由してインポート側でも解析できるよう、
    /// 全言語（JP/EN/CN）のラベルを逆引きできるようにしている。
    /// </summary>
    public static class ParameterKindHelper
    {
        // 接尾辞の囲み記号（パラメータ名にはまず現れない全角記号を使う）
        private const char Open = '【';
        private const char Close = '】';
        private const char IndexSep = '#';

        /// <summary>ライブの Parameter から種別を判定する。</summary>
        public static ParameterKind Determine(Parameter param)
        {
            if (param == null) return ParameterKind.Project;

            try
            {
                if (param.IsShared) return ParameterKind.Shared;
            }
            catch { }

            try
            {
                if (param.Definition is InternalDefinition intDef
                    && intDef.BuiltInParameter != BuiltInParameter.INVALID)
                    return ParameterKind.BuiltIn;
            }
            catch { }

            return ParameterKind.Project;
        }

        /// <summary>種別の表示ラベル（現在の言語）。</summary>
        public static string Label(ParameterKind kind)
        {
            switch (kind)
            {
                case ParameterKind.BuiltIn: return Loc.S("Export.ParamKind.BuiltIn");
                case ParameterKind.Shared: return Loc.S("Export.ParamKind.Shared");
                default: return Loc.S("Export.ParamKind.Project");
            }
        }

        // 言語非依存の逆引き表（Strings{JP,EN,CN}.cs の Export.ParamKind.* と一致させること）。
        // Excel をある言語で書き出し、別言語で取り込んでも解析できるよう全言語を登録する。
        private static readonly Dictionary<string, ParameterKind> ReverseLabels =
            new Dictionary<string, ParameterKind>
            {
                { "組み込み", ParameterKind.BuiltIn },
                { "Built-in", ParameterKind.BuiltIn },
                { "内置", ParameterKind.BuiltIn },
                { "共有", ParameterKind.Shared },
                { "Shared", ParameterKind.Shared },
                { "共享", ParameterKind.Shared },
                { "プロジェクト", ParameterKind.Project },
                { "Project", ParameterKind.Project },
                { "项目", ParameterKind.Project },
            };

        /// <summary>ラベル文字列（全言語対応）から種別を逆引きする。</summary>
        public static bool TryParseLabel(string label, out ParameterKind kind)
        {
            if (!string.IsNullOrEmpty(label) && ReverseLabels.TryGetValue(label.Trim(), out kind))
                return true;
            kind = ParameterKind.Project;
            return false;
        }

        /// <summary>
        /// 同名区別用の接尾辞を生成する。index が 1 以上（同種別で複数ある場合）は
        /// 「#連番」を付ける。例: <c>【組み込み】</c> / <c>【共有#2】</c>。
        /// </summary>
        public static string BuildSuffix(ParameterKind kind, int index)
        {
            string body = index >= 1 ? Label(kind) + IndexSep + index : Label(kind);
            return Open + body + Close;
        }

        /// <summary>
        /// 表示名末尾の曖昧性解消接尾辞（<c>【…】</c>）を解析して取り除く。
        /// 接尾辞が無い／認識できない種別ラベルの場合は false（name はそのまま）。
        /// </summary>
        public static bool TryExtractSuffix(string name, out string baseName, out ParameterKind kind, out int index)
        {
            baseName = name;
            kind = ParameterKind.Project;
            index = 0;

            if (string.IsNullOrEmpty(name) || name[name.Length - 1] != Close)
                return false;

            int open = name.LastIndexOf(Open);
            if (open < 0)
                return false;

            string inner = name.Substring(open + 1, name.Length - open - 2);
            string kindPart = inner;
            int hash = inner.LastIndexOf(IndexSep);
            if (hash >= 0)
            {
                string idxStr = inner.Substring(hash + 1);
                if (int.TryParse(idxStr, out int idx))
                {
                    kindPart = inner.Substring(0, hash);
                    index = idx;
                }
            }

            // 認識できる種別ラベルの時のみ接尾辞とみなす
            // （【】で終わる実パラメータ名を誤って削らないための保険）
            if (!TryParseLabel(kindPart, out kind))
            {
                index = 0;
                return false;
            }

            baseName = name.Substring(0, open);
            return true;
        }
    }
}
