# PABBancolombiaGDocs
(http://codecrastination.com/google-spreadsheet-bancolombia-pab-files/)

### Nomina generadora de archivos PAB para pagos de Bancolombia en Google Sheets

Adquirir software administrativo de nomina, contabilidad, recursos humanos para empresas pequenas de software puede llegar a ser costoso. Este fue uno de los retos que encontramos para una compania de ~25 personas mas pagos de facturas y obligaciones ya que hacer estos pagos de manera manual en la actual sucursal virtual empresarial de bancolombia es demasiado suceptible a errores sin mencionar que es poco amigable la interfaz.

Afortunadamente, Bancolombia [provee un formato en texto plano para transferencias ](http://www.grupobancolombia.com/contenidoCentralizado/corporativo/formatospdf/SVE/FormatoPagosPAB.pdf), de igual manera tambien proveen [una herramienta para crearlos](https://www.satbancolombia.com/FormaArchivos/pab.aspx) pero aun usando estas herramientas, errores pueden suceder al tener que digitar constantemente cuentas bancarias, numeros, fechas, etc. Entonces la solucion seria poner esto en una base de datos, correcto? Pero no queremos instalar todo un motor de bases de datos y mantenerlo y ademas queremos que este en la nube, aqui es donde podemos usar Google Sheets ya que tiene una herramienta de [scripting muy poderoso](https://developers.google.com/apps-script/reference/spreadsheet/?hl=en).

Con esto en mente, en unas pocas horas sacamos adelante el siguiente programita:

Plantilla Google Spreadsheet : https://docs.google.com/spreadsheets/d/1dby-5iSbLo2ZZVMriEx3Fp6H7-froSWl9X-BT29wKUw/edit#gid=127846824

##### Importante: Asegurese que el boton que crea el archivo se le asigne el siguiente script: 
```javascript
CreatePaymentFile()
```
Dudas?
#### oscar@gso.com.co

### Generate Payroll PAB Files for Bancolombia from a Google Sheets Database.

Getting admin software like payroll, accounting, HR for small startups sometimes can be expensive and not the main focus of a software development company. This is a challenge we faced with a payroll of ~25 persons + transfers to vendors and other obligations. Making these payments online manually can be time-consuming and prone to errors.

Fortunately our bank, Bancolombia, [provides a plain text format to create these transactions ](http://www.grupobancolombia.com/contenidoCentralizado/corporativo/formatospdf/SVE/FormatoPagosPAB.pdf), they also provide a [tool to create them](https://www.satbancolombia.com/FormaArchivos/pab.aspx) but even using those tools, errors can happen when typing in bank account details, numbers, dates, etc. So, something like that with a Database would make sense right? But don't want to host anything, maintain a database engine, etc. for a simple program. This is where Google Spreadsheets comes in handy, it allows us to create a simple DB in sheets plus it has got a very powerful scripting tool.

So, in a few hours I came up with this, not very tidy but useful program:

Google Spreadsheet template: https://docs.google.com/spreadsheets/d/1dby-5iSbLo2ZZVMriEx3Fp6H7-froSWl9X-BT29wKUw/edit#gid=127846824

##### Important: Make sure the button to create the file is assigned the following script: 
```javascript
CreatePaymentFile()
```
Queries?
#### oscar@gso.com.co
