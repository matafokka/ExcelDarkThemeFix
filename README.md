# ExcelDarkThemeFix
Fixes Microsoft Excel appearance when custom Windows theme is used

![preview](https://user-images.githubusercontent.com/34414488/90824725-8828aa80-e340-11ea-82db-ef3a5a36ed0e.png)

When the custom dark Windows theme is used, sheets, graphics items, charts and everything that have an "Automatic color" are displayed, exported and printed in the wrong colors.

This happens because Excel picks system colors *(instead of using predefined ones)* for the "Automatic color" which are altered by the themes.

This Add-In aims to fix this issue by setting colors to normal.

# What is fixed?

1. Sheets
1. Cells
1. Shapes
1. Charts - except slightly different font colors
1. Styles
1. Fixing multiple files in one click

# What has workaround?

1. Charts' font colors - there's an option to set it to black.
1. Cell's dark background when editing - there's an option to fix it, but you won't be able to undo anything. As an alternative, use the formula thingy right above the sheet to see what you type.

Enabling these options is described under the "Configuration" section.

# What is not fixed?

1. SmartArt - as far as I've tried, unfixable.
1. Automatic color replacement.
1. Table styles - hard to fix and default styles don't need it.
1. Altering backgrounds - using background is a part of the fix. There's no way of detecting whether the sheet has a background, so scipt just will replace it.
1. You name it, I'm not an Excel expert :D

If you know how to fix any of these things, please, let me know. I will greatly appreciate your effort!

# Compatibility list

Fully compatible with Microsoft Office 2019. Should work in Office 2016 and 360.

If you can test it in versions other than Office 2019, please, let me know, I'll update this list.

# Installation

1. Download `ExcelDarkThemeFix.xlam` and `whitebg.png` files.
1. Navigate to `%AppData%\Microsoft\AddIns` and put the downloaded files here.
1. Enable macros:
   1. Open Excel.
   1. Navigate to "Options".
   1. On the left side of the window click "Customize Ribbon".
   1. On the right side of the window check "Developer" and "Add-Ins" boxes.
   1. Click "Ok" and restart Excel.
1. Enable this Add-In:
   1. Open "Developer" tab.
   1. Click on "Excel Add-Ins".
   1. Check the "ExcelDarkThemeFix" box.
   1. Click "Ok" and restart Excel.

# Configuration

This Add-In provides some cool but unsafe options that you can enable. In order to do this:

1. Open Excel.
1. Open "Developer" tab.
1. Click on "Visual Basic"
1. On the left pane, navigate to element displayed on the following screenshot and double-click on it:

![Open "ExcelDarkThemeFixModule"](https://user-images.githubusercontent.com/34414488/90905218-8ad6de80-e3d8-11ea-8369-9ebb11ffafca.png)

Follow the described instructions. Read everything **carefully**!

After you've done, press `Ctrl+S` and restart Excel.

# FAQ

### Do I need it if I use light theme?

Yes, if your theme uses colors different from white for window and black for text.

### Does it work with a high contrast themes
No.

### Can I disable pop-up dialog when I press "Fix All" button?
Yes, there's an option for this, see "Configuration" section.

### Can you make a dark theme for Excel?
No. There's too much things to track. It's almost (if not entirely) impossible.

### How can I help?

If you're regular user, you can test this Add-In with different kinds of workbooks and report found issues. You could also install an older version of Microsoft Office and check if this Add-In is compatibe with it.

You can also spread the word about this Add-In (especially in non-English speaking parts of the Internet), so other people could have a less of a headache.

If you're developer, you can add some useful features or fix things as described above. Start by thoroughly reading the code to understand what it does. Then do the coding. After you're done, create a pull request with your code.
