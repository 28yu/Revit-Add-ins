using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Tools28.Commands.ExcelExportImport.Models
{
    /// <summary>
    /// パラメータ情報を保持するモデル
    /// </summary>
    public class ParameterInfo : INotifyPropertyChanged
    {
        private bool _isChecked;

        /// <summary>
        /// 表示名（T-/I-プレフィックス＋名前＋曖昧性解消接尾辞）。
        /// 同名パラメータがある場合のみ末尾に種別接尾辞（例: 【組み込み】）が付く。
        /// ダイアログ表示・Excelヘッダー・重複排除キーの単一の情報源。
        /// </summary>
        public string DisplayName => Prefix + RawName + DisambigSuffix;

        /// <summary>プレフィックス（タイプ="T-" / インスタンス="I-"）</summary>
        public string Prefix => IsTypeParameter ? "T-" : "I-";

        /// <summary>
        /// 同名パラメータを区別するための接尾辞（例: 「【組み込み】」）。
        /// 重複が無いパラメータでは空文字（＝従来どおりのヘッダー名）。
        /// </summary>
        public string DisambigSuffix { get; set; } = "";

        /// <summary>パラメータ種別（組み込み/共有/プロジェクト）。同名区別に使用。</summary>
        public ParameterKind Kind { get; set; } = ParameterKind.Project;

        /// <summary>
        /// パラメータの安定 ID（Parameter.Id を long 化）。ドキュメント内で一意なので、
        /// 同名パラメータの厳密な重複排除・照合の識別子として使う。
        /// </summary>
        public long ParamId { get; set; }

        /// <summary>パラメータの元の名前</summary>
        public string RawName { get; set; }

        /// <summary>タイプパラメータの場合true</summary>
        public bool IsTypeParameter { get; set; }

        /// <summary>読み取り専用の場合true</summary>
        public bool IsReadOnly { get; set; }

        /// <summary>要素参照型（StorageType.ElementId）＝文字値ではなく既存要素名が必要</summary>
        public bool IsElementReference { get; set; }

        /// <summary>画像参照パラメータ（ALL_MODEL_IMAGE 等）＝文字値は設定不可</summary>
        public bool IsImage { get; set; }

        /// <summary>所属カテゴリ名</summary>
        public string CategoryName { get; set; }

        /// <summary>
        /// チェック状態（パラメータリスト用）。全選択/選択解除でコードから
        /// 変更した際に CheckBox へ即時反映させるため通知する。
        /// </summary>
        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ParameterInfo(string rawName, bool isTypeParameter, bool isReadOnly, string categoryName)
        {
            RawName = rawName;
            IsTypeParameter = isTypeParameter;
            IsReadOnly = isReadOnly;
            CategoryName = categoryName;
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
                // 同名でも別物のパラメータ（例: エリアの「用途」×2）を区別するため、
                // 表示名ではなく安定 ID（ParamId）＋種別（インスタンス/タイプ）＋カテゴリで同一判定する。
                return ParamId == other.ParamId
                    && IsTypeParameter == other.IsTypeParameter
                    && CategoryName == other.CategoryName;
            }
            return false;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + ParamId.GetHashCode();
                hash = hash * 31 + IsTypeParameter.GetHashCode();
                hash = hash * 31 + (CategoryName?.GetHashCode() ?? 0);
                return hash;
            }
        }
    }
}
