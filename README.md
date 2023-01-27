# PythonOCC-example
This is a program I made that uses PythonOCC to display STL file models that you can interact with in a window.

NOTE: This program uses a faster method to read STL files documented within the SWIG file for PythonOCC.  I could
not actually find any examples of anyone using this faster method.

This program uses PyQt5 backend to display the model files using PythonOCC.  The meat of the use of PythonOCC is contained within
the function labeled "load_model". 

The program does far more than that, as I created it to display a parts heirarchy and keep track of the levels in the heirarchy that you've traveled, with some other features. I removed all references to file locations.

The program itself reads different excel files to populate data for the program, such as populating buttons, names,
parts on hand and etc.  It also uses animations when you click on buttons and travel the "heirarchy" and has a couple additional
features such as loading files when clicking a button.

![1](https://user-images.githubusercontent.com/123666150/215184828-5534fc83-4af8-42c1-adf6-aa4636d56ba7.PNG)
![2](https://user-images.githubusercontent.com/123666150/215184833-a577aba7-2717-400c-8c50-018094a349dd.PNG)
![3](https://user-images.githubusercontent.com/123666150/215184843-1bf9e183-52b9-4ed4-bb56-b5d382d25ebc.PNG)
