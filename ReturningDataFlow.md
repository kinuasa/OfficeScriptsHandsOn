# スクリプトの戻り値を利用するフローのサンプル

- [テーブルを作成するスクリプト](#テーブルを作成するスクリプト)
- [Power Automateから実行するスクリプト](#Power-Automateから実行するスクリプト)
- [フローの作成](#フローの作成)
- [フロー全体図](#フロー全体図)

---

スクリプトからの戻り値を利用する、Power Automateフローの簡単なサンプルです。

## テーブルを作成するスクリプト

まずは使用するシートとテーブルを準備します。  
新規Excelファイルを作成し、下記スクリプトを実行します。

下記スクリプトは、最初のシート名を「**サンプルシート**」とし、セルA1を「**サンプルテーブル**」とするスクリプトです。

```typescript:テーブル作成.ts
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getFirstWorksheet();
  sheet.setName("サンプルシート");
  let range = sheet.getRange("A1");
  range.setValue("サンプル列");
  workbook.addTable(range, true).setName("サンプルテーブル");
}
```

上記スクリプト実行後、「**テーブルに行追加.xlsx**」ファイルとして保存します。

![ReturningDataFlow_01.jpg](images/ReturningDataFlow_01.jpg)

## Power Automateから実行するスクリプト

次はフローから呼び出すスクリプトを準備します。  
「**テーブルに行追加.xlsx**」ファイルを開き、下記スクリプトを「**テーブルの行数取得**」として保存します。

```typescript:テーブルの行数取得.ts
function main(workbook: ExcelScript.Workbook, param: string = "Hello.") {
  let sheet = workbook.getWorksheet("サンプルシート");
  let table = sheet.getTable("サンプルテーブル");
  table.addRow(-1, [param]);
  return table.getRangeBetweenHeaderAndTotal().getRowCount(); //見出し行以外の行数取得
}
```

上記スクリプトは、「**サンプルシート**」上の「**サンプルテーブル**」に行を追加し、見出し行以外の行数を返すスクリプトです。

![ReturningDataFlow_02.jpg](images/ReturningDataFlow_02.jpg)

## フローの作成

次はPower Automateでフローを作成します。

1. Power Automateを開きます。

![ReturningDataFlow_03.jpg](images/ReturningDataFlow_03.jpg)

2. 「作成」から「**インスタント フロー**」をクリックします。

![ReturningDataFlow_04.jpg](images/ReturningDataFlow_04.jpg)

3. フロー名を入力後、「**手動でフローをトリガーします**」を選択し、「**作成**」ボタンをクリックします。

![ReturningDataFlow_05.jpg](images/ReturningDataFlow_05.jpg)

4. 「新しいステップ」から「**スクリプトの実行**」を選択します。

![ReturningDataFlow_06.jpg](images/ReturningDataFlow_06.jpg)

![ReturningDataFlow_07.jpg](images/ReturningDataFlow_07.jpg)

5. 「場所」は「**OneDrive for Business**」、「ドキュメント ライブラリ」は「**OneDrive**」、ファイルは「**テーブルに行追加.xlsx**」、「スクリプト」は「**テーブルの行数取得**」を選択します。

![ReturningDataFlow_08.jpg](images/ReturningDataFlow_08.jpg)

6. スクリプトに渡すパラメーターが設定できるようになるので、「param」として適当な値を指定します(今回は「ユーザー名」)。

![ReturningDataFlow_09.jpg](images/ReturningDataFlow_09.jpg)

7. 「新しいステップ」から「**条件**」を選択します。

![ReturningDataFlow_10.jpg](images/ReturningDataFlow_10.jpg)

8. 条件を「**result**」「**次の値より大きい**」「**3**」とします。

![ReturningDataFlow_11.jpg](images/ReturningDataFlow_11.jpg)

![ReturningDataFlow_12.jpg](images/ReturningDataFlow_12.jpg)

9. 「はいの場合」から「**アクションの追加**」をクリックします。

![ReturningDataFlow_13.jpg](images/ReturningDataFlow_13.jpg)

10. 「**Send me an email notification**」を選択し、「Subject」は「**テーブルの上限を超えました。**」、「Body」は「**テーブルの行数：result**」を指定します。

![ReturningDataFlow_14.jpg](images/ReturningDataFlow_14.jpg)

![ReturningDataFlow_15.jpg](images/ReturningDataFlow_15.jpg)

11. フローを保存し、「**テスト**」ボタンをクリックして動作確認を行います。

![ReturningDataFlow_16.jpg](images/ReturningDataFlow_16.jpg)

![ReturningDataFlow_17.jpg](images/ReturningDataFlow_17.jpg)

![ReturningDataFlow_18.jpg](images/ReturningDataFlow_18.jpg)

![ReturningDataFlow_19.jpg](images/ReturningDataFlow_19.jpg)

12. 「**フローの実行ページ**」を開き、問題無く動作したかどうかを確認します。

![ReturningDataFlow_20.jpg](images/ReturningDataFlow_20.jpg)

![ReturningDataFlow_21.jpg](images/ReturningDataFlow_21.jpg)

![ReturningDataFlow_22.jpg](images/ReturningDataFlow_22.jpg)

13. 問題無くフローが実行されていれば、手順5.で指定したExcelファイルのテーブルに行が追加されていることが確認できます。

![ReturningDataFlow_23.jpg](images/ReturningDataFlow_23.jpg)

14. 数回フローを実行し、手順8.で指定した行数を超えた時点で通知メールが送信されることを確認します。

![ReturningDataFlow_24.jpg](images/ReturningDataFlow_24.jpg)

![ReturningDataFlow_25.jpg](images/ReturningDataFlow_25.jpg)

以上のように、スクリプトからの戻り値を活用することで、フローを分岐させることができます。

## フロー全体図

フローの全体図は下図の通りです。

![ReturningDataFlow_26.jpg](images/ReturningDataFlow_26.jpg)