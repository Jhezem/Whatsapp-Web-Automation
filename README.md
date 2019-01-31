# Whatsapp-Web-Automation - WppAuto
WppAuto fue un proyecto que empecé hace unos meses, pensando en una manera de automatizar el envío de información mediante Whatsapp. Primeramente la idea fue crear un software similar a "Whatsapp Bulk", por lo que busqué una manera de replicar las funcionalidades de éste, en un momento tuve la idea de que podía ser una aplicación de pago, por lo que no publiqué el código abiertamente, sin embargo, al no tener Pc actualmente me he visto en la situación de no poder seguir el proyecto, y para no desperdiciar todo el trabajo realizado, he decidido publicar mi código fuente.

Cabe recalcar, que no soy un experto, pero estoy seguro de que en algo servirá este proyecto, ya sea con código o ideas.

El proyecto fue escrito en C# + Selenium

 ---------------------------------------------------------------------------------------
 
<h2>Uso</h2>
                    
 Para el funcionamiento de la aplicación, es necesario:
 
 1) La aplicación solo funciona con el archivo excel en el proyecto (DB-MENSAJES.Xlm), este archivo excel contiene un código Vba que es esencial para que la aplicación corra, por lo tanto es necesario activar las macros en tu excel. - Este código podría ser fácilmente escrito en C#, así que podrías mejorarlo.
 
 2) Para que la aplicación funcione es necesario tener un chat donde se enviarán los links de la api click to chat.
 
 ![Chat Links](https://drive.google.com/file/d/1iGNZF6bQpX10mdjnlXQsC2kYPe45YpyC/view)
 
 <b>driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[2]/div/label/input")).SendKeys("Nombre del chat" + OpenQA.Selenium.Keys.Enter);</b>
 
 Mira mi video
 
 [![Whatsapp Automation](https://img.youtube.com/vi/CDvbx--OfpI/0.jpg)](https://www.youtube.com/watch?v=YCDvbx--OfpI)
 
<h2>Funcionalidades</h2>

 1) Enviar mensajes a contactos registrados, con lista de nombres.
 
 [![Whatsapp Automation](https://img.youtube.com/vi/3JttBmQHLQo/0.jpg)](https://www.youtube.com/watch?v=3JttBmQHLQo)
 
 2) Enviar mensajes a números sin registrar.
 
 [![Whatsapp Automation](https://img.youtube.com/vi/CDvbx--OfpI/0.jpg)](https://www.youtube.com/watch?v=YCDvbx--OfpI)
 
 3) Enviar imágenes
 
  [![Whatsapp Automation](https://img.youtube.com/vi/wHnEWmsearI/0.jpg)](https://www.youtube.com/watch?v=wHnEWmsearI)
  
  <h1>El uso de este contenido queda bajo responsabilidad de el que lo emplee </h1>
