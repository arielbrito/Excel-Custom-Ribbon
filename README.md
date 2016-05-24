# Ribbon personalizado Office

Para poder personalizar el Ribbon (Cinta de opciones) en Office (Excel principalmente) necesitamos seguir los siguientes pasos y usar los ficheros de la carpeta "Ribbon":

## Personalización de los archivos para añadir el menú:

1. Los archivos necesarios para crear el menú personalizado están en la carpeta "Ribbon" que hay en donde está este archivo README.

2. Editamos el fichero "_customUI\customUI.xml_" que se encuentra en nuestra ruta "_Ribbon_".

3. El **archivo está comentado e informa cómo debemos modificarlo** para que lo tengamos a nuestro gusto.

4. Una vez finalizadas las modificaciones en este "_customUI.xml_" según nuestras necesidades guardamos los cambios.

5. Volvemos a la carpeta "_Ribbon_" donde veremos las dos carpetas necesarias para modificar nuestro Excel y añadirle el menú personalizado: "__Rels_" y "_customUI_".

## Modificación del propio Excel al que se le va a añadir el menú

1. Cambiamos la extensión del excel que vamos a modificar el menú de "_.xlsx_" a "_.zip_".

2. Abrir el fichero con cualquier editor de archivos comprimidos (sirve el que viene con Windows).

3. Veremos una lista de carpetas y archivos.

4. Arrastramos las carpetas (las que hemos modificado en el paso anterior) "__rels_" y "_customUI_" en la ventana del archivo que acabamos de cambiar la extensión.

5. Cerramos el descompresor.

6. Volvemos a cambiar la extensión del archivo de "_.zip_" a "_.xlsx_"

## Documentación

Todo el proceso está documentado paso a paso y con explicaciones detalladas en mi blog BinaryTopic: <https://binarytopic.com/menus-personalizados-en-exel-vba/>
