# Japan Microsoft 365 Developer Community Day 2021:Office Scripts & Power Automate Hands-on

※本資料は、 [Japan Microsoft 365 Developer Community Day 2021](https://jpm365dev.connpass.com/event/227478/) イベントの「**Office スクリプト (Office Scripts) & Power Automate**」ハンズオンセッションの資料です。  
※ハッシュタグ：[#jpm365dcd](https://twitter.com/search?q=%23jpm365dcd&f=live) , [#OfficeScripts](https://twitter.com/search?q=%23OfficeScripts&f=live)

- [ハンズオンの目的](#ハンズオンの目的)
- [ハンズオンの対象者](#ハンズオンの対象者)
- [タイムテーブル](#タイムテーブル)
- [事前準備](#事前準備)
- [ハンズオン環境](#ハンズオン環境)
- [ハンズオン内容](#ハンズオン内容)
- [もっとハンズオン！](#もっとハンズオン！)
- [参考Webサイト](#参考Webサイト)
- [参考動画](#参考動画)
- [Q&Aサイト・フィードバック先](#Q&Aサイト・フィードバック先)
- [Office アドインのMicrosoft Learnコンテンツ](#Office-アドインのMicrosoft-Learnコンテンツ)

---

## ハンズオンの目的

[Ignite 2019](https://news.microsoft.com/ignite2019/)で、Web版のExcel(Excel on the web)での処理をスクリプトで自動化する機能「**Office スクリプト (Office Scripts)**」が発表され、**2021年5月に一般公開**されました。
コードはTypeScript(JavaScript)で書くことができ、VBAの『マクロの記録』機能のように操作を記録・再生することもできます。

本ハンズオンは、「**Office スクリプト**」の概要と開発方法の学習を目的としています。  
実際にスクリプトを書いて実行し、Power Automateとの連携を体験することで、 **“Office スクリプトでどんなことができるのか”** を学んでいきましょう！

## ハンズオンの対象者

本ハンズオンの対象者は下記の通りです。

- Web版のExcel(Excel on the web)を使ったことがあり、処理の自動化に興味がある。
- JavaScript(TypeScript)でコードを書いたことがある。
- Power Automateを知っている、あるいは使ったことがある。
- VBAマクロを書いたことがある。
- セッションレベル：Level 100- Beginner/Introductory 初学者向け ～ Level 200- Intermediate 初級者向け

## タイムテーブル

|  |  |
|------|-------------|
| ご挨拶～ハンズオン説明 | 5分 |
| Microsoft 365 開発者プログラム説明 | 5分 |
| Office スクリプト概要 | 20～30分 |
| ハンズオン(「Excel on the web で Office スクリプトを記録、編集、作成する」、「Excel on the web で Office スクリプトを使用してブックのデータを読み取る」、「ブックに画像を追加する」) | 20分 |
| 休憩 | 5分 |
| ハンズオン(「手動 Power Automation フローからスクリプトを呼び出す」、「フローからスクリプトにデータを渡す方法とスクリプトからデータを返す方法」、「Office スクリプトとPower Automateで見積書を発行する」、「フローからテーブルをフィルタリングして結果を取得する方法」、「フローからスクリプト経由でワークシート関数を実行する方法」) | 40分 |
| もっとハンズオン！～Q&A | n分 |
| まとめ＆アンケートのお願い | 5分 |

## 事前準備

1. Microsoft 365 開発者プログラムの登録

Office スクリプトを利用するには、下記のサブスクリプションが必要となります。  
当サブスクリプションをお持ちでない方は、「[Microsoft 365 開発者プログラムの登録方法](https://github.com/kinuasa/Setup-M365-DevProgram)」を参考に、「[Microsoft 365 開発者プログラム](https://developer.microsoft.com/ja-jp/microsoft-365/dev-program)」に登録してください。本プログラムに登録することで、開発者用のMicrosoft 365 E5サブスクリプション(25ユーザーライセンス)を**無料**で取得できます。

> - Microsoft 365 Business Standard
> - Microsoft 365 Apps for business
> - Microsoft 365 Apps for enterprise
> - Office 365 E3
> - Office 365 E5
> - Office 365 A3
> - Office 365 A5

[https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057&view=o365-worldwide#before-you-begin](https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057&view=o365-worldwide#before-you-begin) より

2. Office スクリプトの有効化

「[スクリプトの可用性Officeスクリプトの共有を管理する](https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057#manage-availability-of-office-scripts-and-sharing-of-scripts)」を参考に、Microsoft 365 管理センターからOffice スクリプトを有効にしてください。なお、本設定の反映には最大48時間かかる場合があります。  
Office スクリプトが利用できる状態であれば、Web版のExcelを開いた際、「自動化」タブが表示されます。  
([※タブが表示されないときは？](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/troubleshooting?WT.mc_id=M365-MVP-4029057#automate-tab-not-appearing-or-office-scripts-unavailable))

## ハンズオン環境

|  |  |
|------|-------------|
| OS | Windows 10 Pro x64 |
| Office | [Web版のMicrosoft Excel](https://www.office.com/launch/excel) (Excel on the web) |
| Browser | [Microsoft Edge(Chromium版)](https://www.microsoft.com/ja-jp/edge), [Google Chrome](https://www.google.com/chrome/) |

## ハンズオン内容

1. Office スクリプト概要
    1. スライド：[Japan Microsoft 365 Developer Community Day 2021 - Office スクリプトハンズオン](https://www.slideshare.net/kinuasa/japan-microsoft-365-developer-community-day-2021-office)
1. Excel on the web で Office スクリプトを記録、編集、作成する (公式チュートリアル)
    1. [データを追加し、基本スクリプトを記録する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#add-data-and-record-a-basic-script)
    1. [既存のスクリプトを編集する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#edit-an-existing-script)
    1. [テーブルを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#create-a-table)
1. Excel on the web で Office スクリプトを使用してブックのデータを読み取る (公式チュートリアル)
    1. [セルを読み取る](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#read-a-cell)
    1. [セルの値を変更する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#modify-the-value-of-a-cell)
    1. [列の値を変更する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#modify-the-values-of-a-column)
1. ブックに画像を追加する (公式チュートリアル)
    1. [サンプル Excel ファイル](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/samples/add-image-to-workbook?WT.mc_id=M365-MVP-4029057#sample-excel-file)
    1. [サンプル コード: ワークシート間で画像をコピーする](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/samples/add-image-to-workbook?WT.mc_id=M365-MVP-4029057#sample-code-copy-an-image-across-worksheets)
    1. [サンプル コード: URL からブックにイメージを追加する](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/samples/add-image-to-workbook?WT.mc_id=M365-MVP-4029057#sample-code-add-an-image-from-a-url-to-a-workbook)
1. 手動 Power Automation フローからスクリプトを呼び出す (公式チュートリアル)
    1. [ブックを準備する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#prepare-the-workbook)
    1. [オフィス スクリプトを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#create-an-office-script)
    1. [Power Automate を使用して自動化されたワークフローを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#create-an-automated-workflow-with-power-automate)
    1. [Power Automate でスクリプトを実行する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#run-the-script-through-power-automate)
1. [フローからスクリプトにデータを渡す方法とスクリプトからデータを返す方法](/DataTransferInFlows.md)
1. [Office スクリプトとPower Automateで見積書を発行する](/PrepareQuote.md)
1. [フローからテーブルをフィルタリングして結果を取得する方法](/GetFilteredResults.md)
1. [フローからスクリプト経由でワークシート関数を実行する方法](/UsingExcelWorksheetFunction.md)

## もっとハンズオン！

余裕がある方は、是非下記内容にもチャレンジしてみてください！ :smile:

1. [サンプルスクリプト](https://docs.microsoft.com/ja-JP/office/dev/scripts/resources/excel-samples?WT.mc_id=M365-MVP-4029057)
1. [Power Automate フローでマクロ ファイルを使用する](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/macros-power-automate?WT.mc_id=M365-MVP-4029057)
1. [Office ScriptsとPower Automateで備品購入申請書を作成する](/RequisitionSlipFlow.md)
1. [スクリプトの戻り値を利用するフローのサンプル](/ReturningDataFlow.md)
1. [[Office Scripts]任意の場所にあるスクリプトを実行する方法](https://www.ka-net.org/blog/?p=13932)
1. [Office ScriptsとPower Automateで簡単なメールアーカイブを作る方法](https://www.ka-net.org/blog/?p=13077)
1. [Cross Reference and formatting two Excel files using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Cross-Reference-and-formatting-two-Excel-files-using-Office/td-p/728535)
1. [Excel and Microsoft Forms integration using Office Scripts](https://powerusers.microsoft.com/t5/Power-Automate-Cookbook/Excel-and-Microsoft-Forms-integration-using-Office-Scripts/td-p/728183)

## 参考資料

1. [Office Scripts(Office スクリプト)の記事まとめ | 初心者備忘録](https://www.ka-net.org/blog/?p=12733)
1. [Office ScriptによるExcel on the web開発入門 | 著者：掌田 津耶乃, 出版：ラトルズ](https://www.rutles.net/products/detail.php?product_id=882)
1. [Office スクリプト API リファレンス | Microsoft Docs](https://docs.microsoft.com/ja-jp/javascript/api/office-scripts/overview?WT.mc_id=M365-MVP-4029057)
1. [Excel on the web での Office スクリプトのスクリプトの基本事項 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/scripting-fundamentals?WT.mc_id=M365-MVP-4029057)
1. [組み込み JavaScript オブジェクト | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/javascript-objects?WT.mc_id=M365-MVP-4029057)
1. [Officeスクリプト コード エディター環境 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/overview/code-editor-environment?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトと VBA マクロの違い | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/vba-differences?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトと Office アドインの違い | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/add-ins-differences?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトでの外部 API 呼び出しのサポート | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/external-calls?WT.mc_id=M365-MVP-4029057)
1. [Power Automate でスクリプトを実行する | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/power-automate-integration?WT.mc_id=M365-MVP-4029057)
1. [トラブルシューティングの基本 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/troubleshooting?WT.mc_id=M365-MVP-4029057)
1. [スクリプトパフォーマンスの機能強化 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/web-client-performance?WT.mc_id=M365-MVP-4029057)
1. [プラットフォームの制限 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/platform-limits?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトの効果を元に戻す | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/undo?WT.mc_id=M365-MVP-4029057)
1. [Officeスクリプトのサンプルとシナリオ | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/scenarios/sample-scenario-overview?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトでのベスト プラクティス | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/best-practices?WT.mc_id=M365-MVP-4029057)

## 参考動画

1. [Automate spreadsheets with Office Scripts in Microsoft Excel | Events](https://docs.microsoft.com/ja-jp/events/build-2020/int114?WT.mc_id=M365-MVP-4029057)
1. [Microsoft 365 Developer (Office Scripts) | YouTube](https://www.youtube.com/c/Microsoft365Developer/search?query=%22Office%20Scripts%22)
1. [Microsoft 365 Community (Office Scripts) | YouTube](https://www.youtube.com/c/Microsoft365PnPCommunity/search?query=%22Office%20Scripts%22)
1. [Sudhi Ramamurthy | YouTube](https://www.youtube.com/user/s65012r/videos)
1. [Office Scripts with Power Automate | YouTube](https://www.youtube.com/watch?v=1jxXXnxdG9A)
1. [What's new in Office Scripts for Excel on the web | YouTube](https://www.youtube.com/watch?v=94YYO3xiSOI)
1. [What’s cooking with Office Scripts: Getting Started | YouTube](https://www.youtube.com/watch?v=FlWerQobJBM)
1. [Excel Office Scripts: Send Teams meeting invite based on Excel table data | YouTube](https://www.youtube.com/watch?v=HyBdx52NOE8)
1. [Excel Office Scripts: Calculate, create Chart, get Chart & Table image, Email | YouTube](https://www.youtube.com/watch?v=152GJyqc-Kw)
1. [Excel Office Scripts: Manage calculate mode, calculate | YouTube](https://www.youtube.com/watch?v=iw6O8QH01CI)
1. [Excel Office Scripts: Use Filter on Table and get Visible Range as Objects | YouTube](https://www.youtube.com/watch?v=Mv7BrvPq84A)
1. [Excel Office Scripts: Clear Hyperlinks from Excel Cells | YouTube](https://www.youtube.com/watch?v=v20fdinxpHU)
1. [Excel Office Scripts: Add comments to Excel Cells | YouTube](https://www.youtube.com/watch?v=CpR78nkaOFw)
1. [Excel Office Scripts: Combine Excel tables into a master table | YouTube](https://www.youtube.com/watch?v=di-8JukK3Lc)
1. [Excel Office Scripts: Move Rows Across Tables and Manage Filters | YouTube](https://www.youtube.com/watch?v=_3t3Pk4i2L0)
1. [Excel Office Scripts: Range basics | YouTube](https://www.youtube.com/watch?v=4emjkOFdLBA)
1. [Excel Office Scripts: Range read and write in perf optimized way (small data) | YouTube](https://www.youtube.com/watch?v=lsR_GvVW3Pg)
1. [Excel Office Scripts: Application basics and environment | YouTube](https://www.youtube.com/watch?v=vvCtxsjPxo8)
1. [Office Scripts: Update large Excel range in performant way | YouTube](https://www.youtube.com/watch?v=BP9Kp0Ltj7U)
1. [API Call from Office Scripts | YouTube](https://www.youtube.com/watch?v=fulP29J418E)
1. [Office Scripts: Add Row at End of Worksheet | YouTube](https://www.youtube.com/watch?v=RgtUar013D0)
1. [Office Scripts: Introduction to the make-up of a script | YouTube](https://www.youtube.com/watch?v=8Zsrc1uaiiU)
1. [Office Scripts: Run Scripts for all Excel files in a folder using Power Automate | YouTube](https://www.youtube.com/watch?v=xMg711o7k6w)
1. [Office Scripts: Top 5 tips to improve your scripting skills in Excel | YouTube](https://www.youtube.com/watch?v=xm2z_D8eP_o)

## Q&Aサイト・フィードバック先

1. [Stack Overflow - office-scripts](https://stackoverflow.com/questions/tagged/office-scripts)
1. [Microsoft Q&A - office-scripts-excel-dev](https://docs.microsoft.com/en-us/answers/topics/office-scripts-excel-dev.html)
1. [Microsoft User Research - Office Scripts Makers](https://ux.microsoft.com/Panel/OfficeScriptsTrade)
1. [Microsoft Feedback Portal - Excel](https://feedbackportal.microsoft.com/feedback/forum/c23f3b77-f01b-ec11-b6e7-0022481f8472)

## Office アドインのMicrosoft Learnコンテンツ

Office スクリプトの兄弟的機能「**Office アドイン**」は、[Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/)で学習コンテンツが用意されています。  
興味がある方は是非こちらもチャレンジしてみてください。

1. [アドインを使用した Office クライアントのカスタマイズの概要 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/intro-office-add-ins/?WT.mc_id=M365-MVP-4029057)
1. [Office アドインで Office クライアントを拡張する | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/paths/m365-office-add-in-associate/?WT.mc_id=M365-MVP-4029057)
1. [Excel 用 Office アドインを作成する | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-excel/?WT.mc_id=M365-MVP-4029057)
1. [Word 用 Office アドインの構築 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-word/?WT.mc_id=M365-MVP-4029057)
1. [Outlook 用 Office アドインの構築 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-outlook/?WT.mc_id=M365-MVP-4029057)