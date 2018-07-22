[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

# エラー時に処理を停止する
$ErrorActionPreference = "Stop"

# SharePoint Online の URL
$siteUrl = 'https://<tenant>.sharepoint.com/sites/<site name>'
# ユーザー名
$user = 'miyamiya@example.com';

try {
    # パスワード
    $secure = Read-Host -Prompt "Enter the password for ${user}(Office365)" -AsSecureString;
    # SharePoint Online 認証情報
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($user, $secure);
    # SharePoint Client Context インスタンスを生成
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
    $ctx.Credentials = $credentials
} catch {
    # SharePoint に接続できない時、エラーログを出して処理を終了する
    Write-Error ($_.Exception.ToString())
    $ctx.Dispose()
    exit
}

# 前回実行時間ファイル
[string]$ProcDatetimeFile = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'SPSync.datetime')

# 前回の処理開始時間
$ProcessedDatetime = (Get-Date).AddDays(-1) # デフォルト 1 日前
if (Test-Path -Path $ProcDatetimeFile) {
    $ProcessedDatetime = [datetime](Get-Content -Path $ProcDatetimeFile).Trim()
}

# 処理開始時間
$procstart = Get-Date

Write-Host ("前回実行時間: {0}" -f $ProcessedDatetime)

# 対象リスト情報を読み込む
$list = $ctx.Web.Lists.GetByTitle('SyncBase')
$ctx.Load($list)
$ctx.ExecuteQuery()

# ChangeQueryの定義
$cq = [Microsoft.SharePoint.Client.ChangeQuery]::new($false, $false)
$cq.Item = $true          # アイテムの更新が対象
$cq.Add = $true           # 追加したアイテム情報が対象
$cq.Update = $true        # 修正したアイテム情報が対象
$cq.DeleteObject = $true  # 削除したアイテム情報が対象
$cq.FetchLimit = 100

# ChangeToken オブジェクトの生成
$cq.ChangeTokenStart = [Microsoft.Sharepoint.Client.ChangeToken]::new()

# ChangeTokenStart に前回実行時間をセット
$cq.ChangeTokenStart.StringValue = ("1;3;{0};{1};-1" -f $list.Id.ToString(), ($ProcessedDatetime.ToUniversalTime().Ticks.ToString()))

try {
    
    # 永久ループ、リミッターを付けるために for 文でも良い
    while(1)
    {
        # 対象リストから更新のあったアイテムを取得
        $changeItems = $list.GetChanges($cq)
        $ctx.Load($changeItems)
        $ctx.ExecuteQuery()

        Write-Host ("Change Item Count: {0}" -f $changeItems.Count.ToString())

        # 前回実行時から変更されたアイテムがない
        if ($changeItems.Count -eq 0) {
            break
        }

        foreach($item in $changeItems){
            switch ($item.ChangeType) {
                "Add" {
                    Write-Host ("ChangeType Add.  ItemId : {0}" -f $item.ItemId)
                }
                "Update" {
                    Write-Host ("ChangeType Update. ItemId : {0}" -f $item.ItemId)
                }
                "DeleteObject" {
                    Write-Host ("ChangeType Delete. ItemId : {0}" -f $item.ItemId)
                }
            }
            # 最後のトークンを取得する
            $lastToken = ([Microsoft.SharePoint.Client.Change]$item).ChangeToken.StringValue
        }
        # 取得した最後のアイテムの ChangeToken.StringValue をセット
        $cq.ChangeTokenStart.StringValue = $lastToken
    }
    # 処理が正常終了したら処理開始時間をファイルに書き込む
    $procstart.ToString('yyyy/MM/dd HH:mm:ss') | Out-File -FilePath $ProcDatetimeFile

} catch {
    # 例外が出たらエラーログを出す
    Write-Error ($_.Exception.ToString())

} finally {
    # Context を破棄
    $ctx.Dispose()

}

