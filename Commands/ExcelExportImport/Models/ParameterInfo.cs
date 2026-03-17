namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// パラメータ情報を保持するモデル
    /// </summary>
    public class ParameterInfo
    {
        /// <summary>パラメータ名（T-/I-プレフィックス付き）</summary>
        public string DisplayName { get; set; }

        /// <summary>パラメータの元の名前</summary>
        public string RawName { get; set; }

        /// <summary>タイプパラメータの場合true</summary>
        public bool IsTypeParameter { get; set; }

        /// <summary>読み取り専用の場合true</summary>
        public bool IsReadOnly { get; set; }

        /// <summary>所属カテゴリ名</summary>
        public string CategoryName { get; set; }

        /// <summary>チェック状態（パラメータリスト用）</summary>
        public bool IsChecked { get; set; }

        public ParameterInfo(string rawName, bool isTypeParameter, bool isReadOnly, string categoryName)
        {
            RawName = rawName;
            IsTypeParameter = isTypeParameter;
            IsReadOnly = isReadOnly;
            CategoryName = categoryName;
            DisplayName = (isTypeParameter ? "T-" : "I-") + rawName;
            IsChecked = false;
        }

        public override string ToString()
        {
            return DisplayName;
        }

        public override bool Equals(object obj)
        {
            if (obj is ParameterInfo other)
            {
                return DisplayName == other.DisplayName && CategoryName == other.CategoryName;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return (DisplayName + CategoryName).GetHashCode();
        }
    }
}
