# sakura2-keisen
サクラエディタ2向けの罫線マクロです。

もともとサクラエディタ(v1)向けの罫線マクロが、[罫線を引く - 上下左右の罫線を引く](http://zenu.xrea.jp/XA5B5A5AFA5E9A5A8A5C7A5A3A5BF2FA5DEA5AFA5ED2FB7D3C0FEA4F2B0FAA4AFX.xhtml)
で公開されていました。

しかし、そのマクロがサクラエディタ2で使用するとまともに動作しなかったため、フォークさせていただく形でサクラエディタ2向けに改変しました。

まだまだ挙動に怪しい部分がありますが、それなりに動作するようになったため、公開することにしました。

## 使い方

* 細線
    * bottom_line.vbs
    * left_line.vbs
    * right_line.vbs
    * top_line.vbs
* 太線
    * bottom_line_b.vbs
    * left_line_b.vbs
    * right_line_b.vbs
    * top_line_b.vbs

細線マクロと太線マクロがそれぞれ上下左右キー分存在しています。

これらを任意のキーにマクロ登録すれば使用できます。

例えば、私は[Alt]+上下左右キーに細線、[Shift]+[Alt]+上下左右キーに太線を登録しています。
