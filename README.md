# Exporting C++ code to Excel : a quick and painless tutorial by Antoine Savine

To export calculation C++ code to Excel is extremely practical. Excel, with all its flexibility and convenience, basically
becomes the front end of your C++ application for free. 
Unfortunately, this is not a particularly intuitive and painless process. 
Books have been written on the topic, including Dalton’s excellent 600 pager:
Financial Applications using Excel Development

For the benefit of my students and colleagues, I wrote a quick and painless tutorial. The tutorial is in the pdf file,
and the rest are the necessary files to make it work. 

I have had excellent feedback from people who found the tutorial particularly helpful, and decided to publish it for everybody’s benefit.
Comments and suggestions are most welcome.

Antoine Savine

### Note added: for 64-bit excel

The tutorial works with 32-bit excel. Nowadays, excel installations often default to 64-bit. If this applies to you, the project decribed in the tutorial will compile and build an xll, but no functionality will be available when you add it into excel. The following steps enabled one user to created a project that works with 64-bit excel:

1.	Download and install the excel sdk from here https://www.microsoft.com/en-us/download/details.aspx?id=35567
    * As part of the installation, you will choose a directory to put it in
2.	Create a new totally separate directory to contain your project, and copy the following files from the sdk directory structure directly to the root of your folder 
    * SRC/XLCALL.CPP
    * INCLUDE/XLCALL.H
    * LIB/x64/XLCALL32.LIB (note from the x64 folder if you are compiling for excel 64bit)
    * SAMPLES/FRAMEWORK/FRAMEWORK.H
    * SAMPLES/FRAMEWORK/FRAMEWORK.C
    * SAMPLES/FRAMEWORK/MemoryManager.cpp
    * SAMPLES/FRAMEWORK/MemoryManager.h
    * SAMPLES/FRAMEWORK/MemoryPool.cpp
    * SAMPLES/FRAMEWORK/MemoryPool.h
3.	From this tutorial, copy the following files into the same directory
    * gaussians.h
    * xlExport.cpp
    * xlOper.h
4.	Follow tutorial as is from now on, with the following notes
    * Note 1, target extension is under Configuration properties / Advanced in more recent versions of visual studio
    * Note 2, when you create the project, precompiled headers will have been automatically created. You don’t need to delete them, but do switch off precompiled headers as described in the tutorial (they just add an unnecessary layer of pain)
    * Add all the above files, including XLCALL32.LIB into your project so it will link against this 64bit version of XLCAL32.LIB
5.	You can build the project in x64 mode
6.	When you build the solution, you will get compile errors like Cannot open include file: 'xlcall.h'. To fix them:
    * Replace statements like #include <xlcall.h> (which fail to compile) with #include “xlcall.h”
    * There are a number of header includes like this that you will need to fix
7.	You will then get compile errors about Excel12 already defined. To fix them:
    * In FRAMEWRK.C, delete/comment out line 76: #include "xlcall.cpp" -> //#include "xlcall.cpp"
8.	In xlOper.h, change #include "framework.h" to become #include "FRAMEWRK.H" 
    * (Alternatively, change the name of the file FRAMEWRK.H to framework.h)

