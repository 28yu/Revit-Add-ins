#!/usr/bin/env bash
# アドイン全体（Tools28.csproj と同じ .cs 一式）を、Linux の dotnet で Revit の各版の API に対してコンパイルする。
# 画面（.xaml）は gen_xaml_stub.py で部品の定義だけを作って補う。
#
# 使い方（リポジトリのルートで）:
#   scripts/compile-check/run.sh            # 2021〜2026 すべて
#   scripts/compile-check/run.sh 2022 2024  # 指定した版だけ
#
# 必要なもの: dotnet SDK 8（Ubuntu なら apt-get install dotnet-sdk-8.0）、nuget.org への接続
# 作業用フォルダは ${COMPILE_CHECK_DIR:-/tmp/tools28-compile-check}（リポジトリの外。消してよい）
set -u
REPO="$(cd "$(dirname "$0")/../.." && pwd)"
WORK="${COMPILE_CHECK_DIR:-/tmp/tools28-compile-check}"
VERSIONS=("$@"); [ ${#VERSIONS[@]} -eq 0 ] && VERSIONS=(2021 2022 2023 2024 2025 2026)

rm -rf "$WORK/gen"; mkdir -p "$WORK/gen"
while IFS= read -r x; do
  python3 -I "$REPO/scripts/compile-check/gen_xaml_stub.py" "$x" "$WORK/gen/$(echo "${x#$REPO/}" | tr '/' '_').g.cs"
done < <(find "$REPO" -name '*.xaml' -not -path "$REPO/Renderer/*" -not -path "*/bin/*" -not -path "*/obj/*")

cat > "$WORK/check.csproj" <<XML
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework Condition="'\$(RevitVersion)' &lt; '2025'">net48</TargetFramework>
    <TargetFramework Condition="'\$(RevitVersion)' &gt;= '2025'">net8.0-windows</TargetFramework>
    <EnableWindowsTargeting>true</EnableWindowsTargeting>
    <EnableDefaultCompileItems>false</EnableDefaultCompileItems>
    <DefineConstants>\$(DefineConstants);REVIT\$(RevitVersion)</DefineConstants>
    <Nullable>disable</Nullable>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>
    <AssemblyName>Tools28</AssemblyName>
    <NoWarn>CS0618;CS1591;CS0649;CS0169;CS0414;CS0067</NoWarn>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Nice3point.Revit.Api.RevitAPI" Version="\$(RevitVersion).*" />
    <PackageReference Include="Nice3point.Revit.Api.RevitAPIUI" Version="\$(RevitVersion).*" />
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3" Condition="'\$(TargetFramework)'=='net48'" />
    <Reference Include="PresentationFramework;PresentationCore;WindowsBase;System.Xaml;System.Windows.Forms;System.Drawing;System.Net.Http;System.IO.Compression;System.IO.Compression.FileSystem;Microsoft.CSharp" Condition="'\$(TargetFramework)'=='net48'" />
    <FrameworkReference Include="Microsoft.WindowsDesktop.App.WPF;Microsoft.WindowsDesktop.App.WindowsForms" Condition="'\$(TargetFramework)'!='net48'" />
    <Compile Include="gen/*.cs" />
    <Compile Include="$REPO/**/*.cs" Exclude="$REPO/Renderer/**;$REPO/**/bin/**;$REPO/**/obj/**" />
  </ItemGroup>
</Project>
XML

# 本体の csproj にある PackageReference（Revit API 以外）も足す
python3 -I - "$REPO/Tools28.csproj" "$WORK/check.csproj" <<'PY'
import re, sys
src, dst = sys.argv[1], sys.argv[2]
s = open(src, encoding='utf-8-sig').read(); d = open(dst, encoding='utf-8').read()
refs = [m.group(0) for m in re.finditer(r'<PackageReference\s+Include="(?!Nice3point)[^"]+"[^>]*?/>', s)]
d = d.replace('    <Compile Include="gen/*.cs" />', '\n'.join('    ' + r for r in refs) + '\n    <Compile Include="gen/*.cs" />')
open(dst, 'w', encoding='utf-8').write(d)
PY

status=0
for v in "${VERSIONS[@]}"; do
  out=$(cd "$WORK" && dotnet build -nologo -v q -p:RevitVersion="$v" -o "out$v" 2>&1)
  errs=$(echo "$out" | grep -E ' error ' | sed "s#$REPO/##; s#\[.*##" | sort -u)
  if [ -n "$errs" ]; then echo "Revit $v: 失敗"; echo "$errs" | head -30; status=1; else echo "Revit $v: OK"; fi
done
exit $status
