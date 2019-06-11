# Limitation

フォームの2回目以降の起動時は、1回目に起動した時にActiveになっていたBookが全面に表示されます  
(フォームから設定する内容は起動の直前にActiveにしていたBookに対する設定)    
理由は、FormObject がメモリ上に生きている為。Office の仕様というか、バグというか。※  
鬱陶しい場合は、一度フォームを右上 `x`ボタン押下か、`Alt + F4` でウィンドウを閉じると、FormOjbect を開放します。  


※  
Excel2013では、UserFormを別のブックから表示させるとActiveなブックが切り変わってしまう。  
[https://answers.microsoft.com/ja-jp/msoffice/forum/all/excel2013%E3%81%A7%E3%81%AFuserform%E3%82%92/0a9b304b-309c-4b8a-96a5-180b9ba9c93c](https://answers.microsoft.com/ja-jp/msoffice/forum/all/excel2013%E3%81%A7%E3%81%AFuserform%E3%82%92/0a9b304b-309c-4b8a-96a5-180b9ba9c93c)

	

