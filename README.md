# Limitation

# Note

Form を Unload せずに再度 show した時は、1回目に show した時の ActiveWorkBook が全面に表示される。  
なのに、VBA から観測する ActiveWorkBook は show する 直前の ActiveWorkBook になる。※  
Excel の仕様というか、バグというか。  

なので、動作終了時は Form を Unload する。

※  
Excel2013では、UserFormを別のブックから表示させるとActiveなブックが切り変わってしまう。  
[https://answers.microsoft.com/ja-jp/msoffice/forum/all/excel2013%E3%81%A7%E3%81%AFuserform%E3%82%92/0a9b304b-309c-4b8a-96a5-180b9ba9c93c](https://answers.microsoft.com/ja-jp/msoffice/forum/all/excel2013%E3%81%A7%E3%81%AFuserform%E3%82%92/0a9b304b-309c-4b8a-96a5-180b9ba9c93c)  
