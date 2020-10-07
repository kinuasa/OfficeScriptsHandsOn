# Global Microsoft 365 Developer Bootcamp 2020:Office Scripts & Power Automate Hands-on

※本資料は、 [Global Microsoft 365 Developer Bootcamp 2020 Tokyo](https://connpass.com/event/188084/) イベントの「**Office Scripts (Office スクリプト) & Power Automate**」ハンズオンセッションの資料です。  
※2020年10月時点では、「Office Scripts」はまだパブリックプレビューです。今後仕様が変更される可能性がありますので、その点はご注意ください。

## ハンズオンの目的

[Ignite 2019](https://news.microsoft.com/ignite2019/)で、Web版のExcel(Excel on the web)での処理をスクリプトで自動化する機能「**Office Scripts (Office スクリプト)**」が発表されました。
コードはTypeScript(JavaScript)で書くことができ、VBAの『マクロの記録』機能のように操作を記録・再生することもできます。

本ハンズオンは、「**Office Scripts**」の概要と開発方法の学習を目的としています。  
実際にスクリプトを書いて実行し、Power Automateとの連携を体験することで、 **“Office Scriptsでどんなことができるのか”** を学んでいきましょう！

## 事前準備

1. Microsoft 365 開発者プログラムの登録

Office Scriptsを利用するには、下記のサブスクリプションが必要となります。  
当サブスクリプションをお持ちでない方は、「[Microsoft 365 開発者プログラムの登録方法](https://github.com/kinuasa/Setup-M365-DevProgram)」を参考に、「[Microsoft 365 開発者プログラム](https://developer.microsoft.com/ja-jp/microsoft-365/dev-program)」に登録してください。本プログラムに登録することで、開発者用のMicrosoft 365 E5サブスクリプション(25ユーザーライセンス)を**無料**で取得できます。

> - Microsoft 365 Business Standard
> - Microsoft 365 Apps for business
> - エンタープライズ向け Microsoft 365 アプリ
> - Office 365 E3
> - Office 365 E5
> - Office 365 A3
> - Office 365 A5

[https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057&view=o365-worldwide#before-you-begin](https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057&view=o365-worldwide#before-you-begin) より

2. Office Scriptsの有効化

「[Office スクリプトの可用性およびスクリプトの共有を管理する](https://docs.microsoft.com/ja-jp/microsoft-365/admin/manage/manage-office-scripts-settings?WT.mc_id=M365-MVP-4029057#manage-availability-of-office-scripts-and-sharing-of-scripts)」を参考に、Microsoft 365 管理センターからOffice Scriptsを有効にしてください。なお、本設定の反映には最大48時間かかる場合があります。  
Office Scriptsが利用できる状態であれば、Web版のExcelを開いた際、「自動化」タブが表示されます。

## ハンズオン環境

|  |  |
|------|-------------|
| OS | Windows 10 Pro x64 |
| Office | [Web版のMicrosoft Excel](https://www.office.com/launch/excel) (Excel on the web) |
| Browser | [Microsoft Edge(Chromium版)](https://www.microsoft.com/ja-jp/edge), [Google Chrome](https://www.google.com/chrome/) |

## ハンズオン内容

1. Office Scripts概要
    1. ※ 「[はじめてのOffice Scripts](https://www.slideshare.net/kinuasa/office-scripts)」をベースにスライド作成予定
1. Excel on the web で Office スクリプトを記録、編集、作成する (公式チュートリアル)
    1. [データを追加し、基本スクリプトを記録する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#add-data-and-record-a-basic-script)
    1. [既存のスクリプトを編集する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#edit-an-existing-script)
    1. [テーブルを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-tutorial?WT.mc_id=M365-MVP-4029057#create-a-table)
1. Excel on the web で Office スクリプトを使用してブックのデータを読み取る (公式チュートリアル)
    1. [セルを読み取る](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#read-a-cell)
    1. [セルの値を変更する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#modify-the-value-of-a-cell)
    1. [列の値を変更する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-read-tutorial?WT.mc_id=M365-MVP-4029057#modify-the-values-of-a-column)
1. 手動 Power Automation フローからスクリプトを呼び出す (公式チュートリアル)
    1. [ブックを準備する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#prepare-the-workbook)
    1. [オフィス スクリプトを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#create-an-office-script)
    1. [Power Automate を使用して自動化されたワークフローを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#create-an-automated-workflow-with-power-automate)
    1. [Power Automate でスクリプトを実行する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-manual?WT.mc_id=M365-MVP-4029057#run-the-script-through-power-automate)
1. 自動で実行される Power Automate フロー内で、データをスクリプトに渡す (公式チュートリアル)
    1. [ブックを準備する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-trigger?WT.mc_id=M365-MVP-4029057#prepare-the-workbook)
    1. [Office スクリプトを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-trigger?WT.mc_id=M365-MVP-4029057#create-an-office-script)
    1. [Power Automate を使用して自動化されたワークフローを作成する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-trigger?WT.mc_id=M365-MVP-4029057#create-an-automated-workflow-with-power-automate)
    1. [Power Automate でスクリプトを管理する](https://docs.microsoft.com/ja-JP/office/dev/scripts/tutorials/excel-power-automate-trigger?WT.mc_id=M365-MVP-4029057#manage-the-script-in-power-automate) 

## もっとハンズオン！

余裕がある方は、是非下記内容にもチャレンジしてみてください！ :smile:

1. [サンプルスクリプト](https://docs.microsoft.com/ja-JP/office/dev/scripts/resources/excel-samples?WT.mc_id=M365-MVP-4029057)
1. ※Power Automateとの連携を中心に課題作成予定

## 参考Webサイト

1. [Office Scripts(Office スクリプト)の記事まとめ | 初心者備忘録](https://www.ka-net.org/blog/?p=12733)
1. [Office スクリプト API リファレンス | Microsoft Docs](https://docs.microsoft.com/ja-jp/javascript/api/office-scripts/overview?WT.mc_id=M365-MVP-4029057)
1. [Excel on the web での Office スクリプトのスクリプトの基本事項 (プレビュー) | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/scripting-fundamentals?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトでの組み込みの JavaScript オブジェクトの使用 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/javascript-objects?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトのコードエディター環境 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/overview/code-editor-environment?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトと VBA マクロの相違点 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/vba-differences?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトと Office アドインの違い | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/add-ins-differences?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトでの外部 API 呼び出しのサポート | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/external-calls?WT.mc_id=M365-MVP-4029057)
1. [Power Automate でスクリプトを実行する | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/power-automate-integration?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトのトラブルシューティング | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/troubleshooting?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトのパフォーマンスを向上させる | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/web-client-performance?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトを使用したプラットフォームの制限と要件 | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/platform-limits?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトを実行して行った変更を元に戻す | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/testing/undo?WT.mc_id=M365-MVP-4029057)
1. [非同期 Api を使用する古い Office スクリプトをサポートする | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/develop/excel-async-model?WT.mc_id=M365-MVP-4029057)
1. [Office スクリプトのサンプルシナリオ | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/dev/scripts/resources/scenarios/sample-scenario-overview?WT.mc_id=M365-MVP-4029057)
1. [Automate spreadsheets with Office Scripts in Microsoft Excel | Channel 9](https://channel9.msdn.com/events/Build/2020/INT114?WT.mc_id=M365-MVP-4029057)

## Office アドインのMicrosoft Learnコンテンツ

Office Scriptsの兄弟的機能「**Office アドイン**」は、[Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/)で学習コンテンツが用意されています。  
興味がある方は、是非こちらもチャレンジしてみてください。

1. [アドインを使用した Office クライアントのカスタマイズの概要 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/intro-office-add-ins/?WT.mc_id=M365-MVP-4029057)
1. [Office アドインで Office クライアントを拡張する | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/paths/m365-office-add-in-associate/?WT.mc_id=M365-MVP-4029057)
1. [Excel 用 Office アドインを作成する | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-excel/?WT.mc_id=M365-MVP-4029057)
1. [Word 用 Office アドインの構築 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-word/?WT.mc_id=M365-MVP-4029057)
1. [Outlook 用 Office アドインの構築 | Microsoft Learn](https://docs.microsoft.com/ja-jp/learn/modules/office-add-ins-outlook/?WT.mc_id=M365-MVP-4029057)