# Ribbon personalizado Office

Para poder personalizar el Ribbon (Cinta de opciones) en Office (Excel principalmente) necesitamos seguir los siguientes pasos y usar los ficheros adjuntos:

## Personalización de los archivos para añadir el menú:

1. Los archivos necesarios para crear el menú personalizado están en la carpeta "Ribbon" que hay en donde está este archivo README.
2. Editamos el fichero "customUI\customUI.xml" que se encuentra en nuestra ruta "Ribbon".
3. El archivo está comentado e informa cómo debemos modificarlo para que lo tengamos a nuestro gusto.
4. Una vez finalizadas las modificaciones en este "customUI.xml" según nuestras necesidades guardamos los cambios.
5. Volvemos a la carpeta "Ribbon" donde veremos las dos carpetas necesarias para modificar nuestro Excel y añadirle el menú personalizado: "_Rels" y "customUI".

## Modificación del propio Excel al que se le va a añadir el menú

1. Cambiamos la extensión del excel que vamos a modificar el menú de ".xlsx" a ".zip".
2. Abrir el fichero con 7zip o cualquier editor de archivos comprimidos (sirve el que viene con Windows).
3. Veremos una lista de carpetas y archivos.
4. Arrastramos las carpetas (las que hemos modificado en el paso anterior) "_rels" y "customUI" en la ventana del archivo que acabamos de cambiar la extensión.
5. Cerramos el descompresor.
6. Volvemos a cambiar la extensión del archivo de ".zip" a ".xlsx"
