jxls-example
============
jXLSのforEachで、行結合した範囲を含んでいると、結合範囲がおかしくなるのを回避する方法のサンプルです。  
`forEachRowsMerge.xls`ファイルには以下の3つのシートが含まれています。
* jxlsBug forEachでの結合の不具合を再現します。
* PoiObjectsAccess PoiObjectsAccessの機能を使ってセル結合した例です。  
* Helper インスタンス化可能なUtilクラスを作ってテンプレートから呼び出す例です。

jUnitで`ForEachTest#testForEach中での行結合`を実行すると、
`tareget\forEachRowsMerge_output.xls`に変換結果を出力します。
