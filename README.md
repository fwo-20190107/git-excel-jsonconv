# git-excel-jsonconv
xls,xlsx形式のファイルをjsonで出力してGit Diffに利用する
元：git-xlsx-textconv

## TODO
- .xlsの場合、文字装飾や式だとセルの文字列を取得できない
- jsonだとどのシートのどのあたりでの差分なのかパッと見でわからない
    - textconvはこのあたりがわかりやすかった
