# ExcelDarkThemeFix
Fixes Microsoft Excel appearance when custom Windows theme is used

![preview](https://user-images.githubusercontent.com/34414488/90824725-8828aa80-e340-11ea-82db-ef3a5a36ed0e.png)

# What is it? What it does?
When the custom dark theme is used, sheets, graphics items, charts and everything that have "Automatic color" are dark. This happens because Excel picks system colors *(instead of using predefined ones)* for "automatic" color, and theme alters it.

As a result, everything is displayed in a wrong color. Moreover, Excel would save and print the document in the wrong colors too.

This Add-In aims to fix this issue by manually setting colors to normal.

# What is fixed?

1. Background
1. Cell colors
1. Shapes
1. Charts
1. Styles

# What is not fixed?

1. SmartArt. It's not possible to fix it.
1. Charts' text color. *Automatic color is close to the needed one, and implementing the fix is hard and mostly useless.*
1. Dark background in a cell when editing text. *Use the formula thingy right above the sheet to see what you type. It's possible to fix, but it will remove the ability to undo things.*
1. Auto color replacement, like when you set the automatic color, it would apply the right one instead. *Hard to fix, will remove the ability to undo things. The fix button also won't be added because of this behavior. If you **really** need to fix everything, save, close and open the workbook again.*
1. You name it, I'm not an Excel expert :D

If you know how to fix any of these things, please, let me know. I will greatly appreciate your effort!

# How to install?
1. Download `ExcelDarkThemeFix.xlam` and `whitebg.png`.
1. Google "excel install add-in" and follow any of the available guides.
1. When you'll move the files, move both files you've downloaded.
