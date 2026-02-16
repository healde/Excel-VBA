# create connection to, and import data from : another workbook

## 🎼 Why ? 
It is one of the things that I wanted to add to my "how to in Vba" skills. But it appears as a less atempted option than I though and the documentation let me go forward through difficulties as the beginer as I am. So I decide to make a full snippet for looking over process and understanding the behaviors of the differents functionality around this subject.
### start
Without finding as much as solutions I expect in forums, I recorded Macro over manual process, and get by chance more than one methods. I cleaned them out of the records context, and tried to make them work as much as I can.
### how then
After having compose a whole set of commands to run some tests, I asked for help to AI and it let me knows about some internal process in the Microsoft environment and followed me in my course.

## 🎶 Basic work
A Query made of a formula coded in (PowerQuery) M langage is called through `ActiveSheet.ListObjects.Add` method to generate a data table from elsewhere. It is possible to get directly the data, or to manage the connection with the distant workbook before getting the data. But in this  last case, documentation is especialy confusing and differences emerge between Vba functions and manual process.

The distant workbook is so described by the M formula, and have to be the same as well as its location must be correct (of course).

## 🎨 Panel
In this program, I purpose 



## 🧩 
