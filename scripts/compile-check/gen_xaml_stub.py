#!/usr/bin/env python3
"""
XAML の x:Class / x:Name を読み、確認用の partial クラス（部品のフィールドと空の InitializeComponent）を作る。

Linux の dotnet には WPF の XAML コンパイル（WindowsDesktop SDK）が無いため、.xaml.cs をコンパイルできない。
このスクリプトで XAML が生成するはずの部分だけを作ると、.xaml.cs も含めて型の検査ができる。
⚠ XAML 自体の誤り（イベントハンドラ名の綴り・属性の誤りなど）は検査できない。

使い方: gen_xaml_stub.py <入力.xaml> <出力.g.cs>
"""
import re
import sys

# x:Name の要素名 → 型の名前空間（ここに無いものは System.Windows.Controls）
NAMESPACES = {
    'ColumnDefinition': 'System.Windows.Controls', 'RowDefinition': 'System.Windows.Controls',
    'Run': 'System.Windows.Documents', 'Hyperlink': 'System.Windows.Documents', 'Span': 'System.Windows.Documents',
    'Bold': 'System.Windows.Documents', 'Italic': 'System.Windows.Documents',
    'Rectangle': 'System.Windows.Shapes', 'Ellipse': 'System.Windows.Shapes', 'Path': 'System.Windows.Shapes',
    'Line': 'System.Windows.Shapes', 'Polygon': 'System.Windows.Shapes', 'Polyline': 'System.Windows.Shapes',
    'SolidColorBrush': 'System.Windows.Media', 'RotateTransform': 'System.Windows.Media',
    'ScaleTransform': 'System.Windows.Media', 'TranslateTransform': 'System.Windows.Media',
    'Storyboard': 'System.Windows.Media.Animation',
    'ToggleButton': 'System.Windows.Controls.Primitives', 'Popup': 'System.Windows.Controls.Primitives',
    'RepeatButton': 'System.Windows.Controls.Primitives', 'Thumb': 'System.Windows.Controls.Primitives',
}


def main(src, out):
    x = open(src, encoding='utf-8-sig').read()
    m = re.search(r'<\s*([A-Za-z]+)\b[^>]*\bx:Class="([^"]+)"', x, re.S)
    root, cls = m.group(1), m.group(2)
    root_ns = 'System.Windows' if root == 'Window' else NAMESPACES.get(root, 'System.Windows.Controls')
    ns, name = cls.rsplit('.', 1)
    fields = []
    for fm in re.finditer(r'<\s*([A-Za-z]+(?::[A-Za-z]+)?)\b[^>]*?\bx:Name="([^"]+)"', x, re.S):
        t, n = fm.group(1), fm.group(2)
        if ':' in t:  # 独自の型（local:Xxx など）は検査しない
            fields.append(f'        internal object {n};')
            continue
        fields.append(f'        internal {NAMESPACES.get(t, "System.Windows.Controls")}.{t} {n};')
    code = (f'namespace {ns}\n{{\n    public partial class {name} : {root_ns}.{root}\n    {{\n'
            + '\n'.join(fields)
            + '\n        public void InitializeComponent() { }\n    }\n}\n')
    open(out, 'w', encoding='utf-8').write(code)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
