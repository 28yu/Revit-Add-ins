using System.Windows.Controls;
using Tools28.Localization;

namespace Tools28.Commands.BeamCommon
{
    /// <summary>
    /// 設定画面の「色分けする梁」（すべて／水平だけ／斜めだけ）のラジオボタン。梁天端・梁下端で共通
    /// </summary>
    public static class BeamSlopeRadios
    {
        /// <summary>本数を文言に入れる。0本の選択肢は選べなくし、選ばれていたら「すべて」に戻す</summary>
        public static void ShowCounts(RadioButton all, RadioButton flat, RadioButton sloped, BeamSlopeCounts counts)
        {
            if (counts == null) return;
            all.Content = string.Format(Loc.S("Beam.Slope.All"), counts.All);
            flat.Content = string.Format(Loc.S("Beam.Slope.Flat"), counts.Flat);
            sloped.Content = string.Format(Loc.S("Beam.Slope.Sloped"), counts.Sloped);
            flat.IsEnabled = counts.Flat > 0;
            sloped.IsEnabled = counts.Sloped > 0;
            if ((flat.IsChecked == true && !flat.IsEnabled) || (sloped.IsChecked == true && !sloped.IsEnabled))
                all.IsChecked = true;
        }

        public static BeamSlopeFilter Read(RadioButton flat, RadioButton sloped)
        {
            if (flat.IsChecked == true) return BeamSlopeFilter.FlatOnly;
            if (sloped.IsChecked == true) return BeamSlopeFilter.SlopedOnly;
            return BeamSlopeFilter.All;
        }

        /// <summary>確認画面の要約に出す文言</summary>
        public static string Describe(BeamSlopeFilter filter)
        {
            switch (filter)
            {
                case BeamSlopeFilter.FlatOnly: return Loc.S("Beam.SlopeName.Flat");
                case BeamSlopeFilter.SlopedOnly: return Loc.S("Beam.SlopeName.Sloped");
                default: return Loc.S("Beam.SlopeName.All");
            }
        }
    }
}
