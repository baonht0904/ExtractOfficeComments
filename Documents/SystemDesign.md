# Requirements

## Functional requirements
+ User can extract comments from Word, Power point and Excel into an Excel file.
+ User can select multiple file in different types with a GUI.
+ User can save the result with a desired name.

## Non-functional requirements
+ The maximun files can be selected depend on the limit of "Open file dialog".

## System constant
+ OS: Win10, Win11
+ Have to support docx, xlsx, pptx; better if support doc, xls, ppt but not required.

# System design

## Basic workflow
![Basic work flow of the Application](/Documents/Image/BasicWorkFlow.svg)

## UI/UX
- UI design:  
![UI of the Application](/Documents/Image/UI.svg)
- Screen items:
    - File exployer view: A list view used to select file from local file system.
    - Selected files view: A list view used to display the selected files.
    - \[<-\] button: A button used to move backward in File exployer view.
    - \[->\] button: A button used to move forward in File exployer view.
    - \[Add\] button: A button used to add files that are selected in File exployer view to Selected files view.
    - \[Remove\] button: A button used to remove files from Selected files view.
    - \[Extract comment\] button: A button used to start the extract process.
- Screen processing:
    - Users use File exployer view, \[<-\], \[->\], \[Add\] and \[Remove\] button to selected the desired files that contain comments.
    - Users click \[Extract comment\] button. A "Save as file" dialog will be appeared for users to select a desired name that is used to contain the result.
    - A progress bar and extracting status will be display in a pop-up screen.
    - After finish extracting, a message will be appeared to notify user.
