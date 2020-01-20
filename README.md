# CustomMsgboxClass

> Una Caja de Mensajes alterantiva para Excel para aplicarse con VBA macros en UserForms.

## Definición
CustomMsgboxClass es un archivo de Clase VBA (.cls) con el que es fácil generar una rápida Caja de Mensajes y habilitarla sólo en Formularios (UserForms) con macros.

> Es necesario tener conocimientos básicos de programación de Clases VBA (POO, Programación Orientada a Objetos).

Trabaja bajo MS Excel versión 2007 o superior. No requiere instalación, sólo debe ser importardo a su Proyecto VBA.

Como CustomMsgboxClass es una Clase VBA, debe instanciarse como un objeto y luego usar su interfaz de métodos y propiedades.

## Modo de uso
  1.  [Descargue el archivo CustomMsgboxClass.cls](https://github.com/Roccouu/CustomMsgboxClass/tree/master/project-dev/CustomMsgboxClass.cls).
  2.  Cree un nuevo libro habilitado para macros de Excel.
  3.  Abra el Editor de Proyectos VBA con **Ctrl+F11**
  4.  Vaya al menú ***Archivo > Importar Archivo*** o presione ***Ctrl+M*** y luego busque su archivo **CustomMsgboxClass.cls** descargado e impórtelo al Proyecto VBA.
  5.  Puede crear un nuevo módulo y un nuevo formulario donde se utilizará ***CustomMsgboxClass*** (también puede descargar [custommsgboxexampleuse.xlsm](https://github.com/Roccouu/CustomMsgboxClass/tree/master/project-dist/custommsgboxexampleuse.xlsm) para ver algunos ejemplos prácticos y fáciles de usar ***CustomMsgboxClass***)
  6.  Llame a uno de los dos métodos del Objeto y...
  7. ¡Disfrute de **CustomMsgboxClass**!

## A cerca de los dos métodos de CustomMsgboxClass
  A. ***El Método CMsgbox:*** Permite crear un Objeto Caja de Mensaje tipo ventana, éste objeto puede mostrarse en un formulario de modo ***asíncrono*** (sín límite de tiempo) o ***síncrono*** (por un tiempo establecido); el modo asíncrono despliega un botón que permite al Usuario cerrar la ventana cuando él lo decida, con el modo síncrono, la ventana se cierra tras haber trasncurrido un tiempo establecido por el desarrollador (por defecto se muestra durante tres segundos y como máximo puede mostrarse durante diez segundos) sín mostrar el botón de cerrar ventana.
  El método tiene varios parámetros, la mayoría opcionales:
  ```vb
    CMsgbox( _
      ByVal MFrm As MSForms.UserForm, _
      Optional MMsg As String = "", _
      Optional MMsgType As String = "", _
      Optional MTitle As String = "CustomMsgboxClass", _
      Optional MSubtitle As String = "", _
      Optional MCloseButton As Boolean = False, _
      Optional MPosition As String = "Middle", _
      Optional MControl As Object, _
      Optional MRequiredControl As Boolean = False, _
      Optional MColor As Long = 4616993, _
      Optional MTime As Single = 3)
  ```
  Donde:

  1.  **Mfrm: (Requerido)** Debe ser un Objeto tipo UserForm donde la Caja de Mensaje se desplegará.
  2.  **MMsg:** Es el mensaje principal que se desea mostrar, es decir, el contenido. El valor por defecto es una cadena vacía.
  3.  **MMsgType:** Es el tipo de Caja de Mensaje que se desea mostrar. Existen cinco tipos posibles:
        -  "Error": Para mostrar mensajes de error.
        -  "Success": Para mostrar mensajes de resultados correctos.
        -  "Info": Para mostrar mensajes de información.
        -  "Question": Para mostrar mensajes de interrogación.
        -  "Alert": Para mostrar mensajes de alerta.
        -  El tipo por defecto sólo utiliza los colores del sistema, sín íconos ni barra de título.
  4.  **MTitle:** El título del mensaje, se ve en la zona de barra de título de la Caja de Mensaje. Valor por defecto: "CustomMsgboxClass".
  5.  **MSubtitle:** Un subtítulo opcional para el mensaje. Cadena vacía como valor por defecto.
  6.  **MCloseButton:** Valor booleano, en verdadero, muestra un botón de cerrar ventana; sólo trabaja en caso ***asíncrono***. Por defecto es Falso.
  7.  **MPosition:** Parámetro tipo cadena. Es la posición del Objeto en el área del Formulario, existen tres posibles posiciones: *"Top"*, muestra el Objeto en la parte superior, *"Middle"*, al medio, y *"Bottom"*, en la parte inferior del Formularrio. Valor por defecto: *"Middle"*.
  8.  **MControl:** Objeto tipo MSForms.Control. En caso de que desee resaltar un control de entrada de datos del Formulario, envíe en *MControl* el Control de su Formulario que desee resaltar (El Objeto trabaja con controles de tipo TextBox, ComboBox, ListBox, RefEdit y Label), *CustomMsgBox* se mostrará justo debajo y alineado a la izquierda del control que asigne (*MPosition* será ignorado), además el control cambiará de color de fondo y borde durante unos segundos, dependiendo del tipo de Caja de Mensaje que haya establecido y de haber confirmado que quiere resaltar el control indicado, luego los estilos de dicho control retornarán a su estado original. Valor por defecto: Nothing.
  9.  **MRequiredControl:** Booleano. Establezca su valor en Verdadero para confirmar que desea resaltar el control enviado en *MControl*. Valor por defecto: Falso.
  10. **MColor:** Tipo de dato Entero Largo (Long). Permite establecer un color de tema para la Caja de Mensaje, puede utilizar un número tipo entero largo ó la functión VBA.RGB(RR,GG,BB) que determine el color de tema que desee para el Objeto. Valor por defecto: 4616993 (VBA.RGB(33, 115, 70), color verde de la Aplicación Excel)
  11. **MTime:** Entero Sencillo. Un valor entre 3 y 10 que permite establecer un tiempo personalizado para mostrar el Obejto Caja de Mensaje, luego de transcurrido ese tiempo, el Objeto se cerrará y autodestruirá del Formulario.

  B. ***El Método CMsgboxFluid:*** Permite crear un Objeto Caja de Mensaje tipo cinta (fluido, expandido en todo el ancho del Formulatio), éste objeto puede mostrarse en un formulario sólo de modo ***síncrono*** (por un tiempo establecido).
  El método tiene varios parámetros, la mayoría opcionales:
  ```vb
    CMsgboxFluid( _
      ByVal MFrm As MSForms.UserForm, _
      Optional MMsg As String = "", _
      Optional MMsgType As String = "", _
      Optional MSubtitle As String = "CustomMsgboxClass", _
      Optional MPosition As String = "Middle", _
      Optional MControl As Object, _
      Optional MRequiredControl As Boolean = False, _
      Optional MColor As Long = 4616993, _
      Optional MTime As Single = 3)
  ```
  Donde:

  1.  **Mfrm: (Requerido)** Debe ser un Objeto tipo UserForm donde la Caja de Mensaje se desplegará.
  2.  **MMsg:** Es el mensaje principal que se desea mostrar, es decir, el contenido. El valor por defecto es una cadena vacía.
  3.  **MMsgType:** Es el tipo de Caja de Mensaje que se desea mostrar. Existen cinco tipos posibles:
      -  "Error": Para mostrar mensajes de error.
      -  "Success": Para mostrar mensajes de resultados correctos.
      -  "Info": Para mostrar mensajes de información.
      -  "Question": Para mostrar mensajes de interrogación.
      -  "Alert": Para mostrar mensajes de alerta.
      -  El tipo por defecto sólo utiliza los colores del sistema, sín íconos ni Subtítulo.
  4.  **MSubtitle:** Un subtítulo opcional para el mensaje. Cadena vacía como valor por defecto.
  5.  **MPosition:** Parámetro tipo cadena. Es la posición del Objeto en el área del Formulario, existen tres posibles posiciones: *"Top"*, muestra el Objeto en la parte superior, *"Middle"*, al medio, y *"Bottom"*, en la parte inferior del Formularrio. Valor por defecto: *"Middle"*.
  6.  **MControl:** Objeto tipo MSForms.Control. En caso de que desee resaltar un control de entrada de datos del Formulario, envíe en *MControl* el Control de su Formulario que desee resaltar (El Objeto trabaja con controles de tipo TextBox, ComboBox, ListBox, RefEdit y Label), *CustomMsgBox* se mostrará justo debajo y alineado a la izquierda del control que asigne (*MPosition* será ignorado), además el control cambiará de color de fondo y borde durante unos segundos, dependiendo del tipo de Caja de Mensaje que haya establecido y de haber confirmado que quiere resaltar el control indicado, luego los estilos de dicho control retornarán a su estado original. Valor por defecto: Nothing.
  7.  **MRequiredControl:** Booleano. Establezca su valor en Verdadero para confirmar que desea resaltar el control enviado en *MControl*. Valor por defecto: Falso.
  8. **MColor:** Tipo de dato Entero Largo (Long). Permite establecer un color de tema para la Caja de Mensaje, puede utilizar un número tipo entero largo ó la functión VBA.RGB(RR,GG,BB) que determine el color de tema que desee para el Objeto. Valor por defecto: 4616993 (VBA.RGB(33, 115, 70), color verde de la Aplicación Excel)
  9. **MTime:** Entero Sencillo. Un valor entre 3 y 10 que permite establecer un tiempo personalizado para mostrar el Obejto Caja de Mensaje, luego de transcurrido ese tiempo, el Objeto se cerrará y autodestruirá del Formulario.

## Colaborar en GitHub:
El código fuente de **CustomMsgboxClass** está en: [el directorio project-dev](https://github.com/Roccouu/CustomMsgboxClass/tree/master/project-dev/CustomMsgboxClass.cls) del repositorio oficial.

Tan pronto como se descargue, puede colaborar con mejoras en el Sistema siempre bajo el respeto de [Términos de licencia](https://github.com/Roccouu/CustomMsgboxClass/blob/master/LICENSE), [El Código de Conducta](https://github.com/Roccouu/CustomMsgboxClass/blob/master/CODE_OF_CONDUCT.md) y los [Términos de Contribución](https://github.com/Roccouu/CustomMsgboxClass/blob/master/CONTRIBUTING.md).

## Sitio Web

[CustomMsgboxClass](https://roccouu.github.io/CustomMsgboxClass/docs/index.html)

## Tutorial

[Tutorial CustomMsgboxClass](https://roccouu.github.io/CustomMsgboxClass/docs/index.html#/tutorial)

## Documentación

[Documentación CustomMsgboxClass](https://roccouu.github.io/CustomMsgboxClass/index.html#/docs/index.html#/documentation)

## Contribución

Vea las [Guías de CONTRIBUCIÓN](https://github.com/roccouu/CustomMsgboxClass/CONTRIBUTING.md)

## English Readme

[README-EN.md](https://github.com/roccouu/CustomMsgboxClass/blob/master/README-EN.md)

## Licencia

[MIT](https://github.com/roccouu/CustomMsgboxClass/blob/master/LICENSE) © | [E-Mail](rocky.romay@gmail.com) | [Roccou](https://twitter.com/_roccou) | 2020