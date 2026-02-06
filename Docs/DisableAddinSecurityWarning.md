# Revit アドインのセキュリティ警告を無効化する方法

## 問題
Revit起動時に「外部ツール: 見つからないアドイン アセンブリ」または「このアドインをロードしますか？」という警告ダイアログが表示される。

## 原因
- アドインにデジタル署名がない
- Revitのセキュリティ設定で未署名のアドインをブロックしている

## 解決方法

### 方法1: Revit設定で「常に読み込む」を選択（最も簡単）

1. Revit起動時に警告ダイアログが表示される
2. **「常に読み込む」**を選択
3. 次回からこのアドインの警告は表示されなくなります

### 方法2: Revitのセキュリティレベルを下げる（非推奨）

⚠️ **注意**: セキュリティリスクがあるため推奨しません。

1. Revitのオプションを開く
2. 「セキュリティ」タブ
3. 「アドインのセキュリティレベル」を**「低」**に設定

### 方法3: アドインにデジタル署名を追加（開発者向け）

#### 前提条件
- Visual Studio 2022
- コード署名証明書（自己署名証明書または商用証明書）

#### 手順

**A. 自己署名証明書の作成**

PowerShellを管理者権限で開き、以下を実行:

```powershell
# 証明書を作成
$cert = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject "CN=Tools28 Development" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter (Get-Date).AddYears(5)

# 証明書のサムプリントを表示
$cert.Thumbprint

# 信頼されたルート証明機関に追加（ローカルPCのみ）
Export-Certificate -Cert $cert -FilePath "C:\temp\Tools28Dev.cer"
Import-Certificate -FilePath "C:\temp\Tools28Dev.cer" -CertStoreLocation "Cert:\CurrentUser\Root"
```

**B. プロジェクトファイルに署名設定を追加**

`Tools28.csproj` に以下を追加:

```xml
<PropertyGroup>
  <SignAssembly>true</SignAssembly>
  <AssemblyOriginatorKeyFile>Tools28.snk</AssemblyOriginatorKeyFile>
</PropertyGroup>
```

**C. ビルド後に署名**

`Tools28.csproj` に以下を追加:

```xml
<Target Name="SignAssembly" AfterTargets="AfterBuild">
  <Exec Command="signtool sign /sha1 証明書のサムプリント /fd SHA256 /t http://timestamp.digicert.com &quot;$(TargetPath)&quot;" />
</Target>
```

**D. ビルド実行**

```powershell
.\BuildAll.ps1
```

## 推奨される方法

開発段階では **方法1**（「常に読み込む」を選択）が最も簡単です。

公開配布する場合は、**方法3**（デジタル署名）を使用してユーザーの信頼を得ることを推奨します。

## 参考
- [Autodesk Revit Add-In Security](https://help.autodesk.com/view/RVT/2022/ENU/?guid=Revit_API_Revit_API_Developers_Guide_Introduction_Add_In_Integration_Add_in_Security_html)
