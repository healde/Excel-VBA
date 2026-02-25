# Create connection to, and import data from : another workbook

## 🎼 Why ? 
It is one of the things that I wanted to add to my "how to in Vba" skills. But it appears as a less atempted option than I though and the documentation let me go forward through difficulties as the beginer as I am. So I decide to make a full snippet for looking over process and understanding the behaviors of the differents functionality around this subject.
### start
Without finding as much as solutions than I expected in forums, I recorded Macro over manual process, and get by chance more than one methods. I cleaned them out of the records context, and tried to make them work as much as I can.
### how then
After having compose a whole set of commands to run some tests, I asked for help to AI and it let me knows about some internal process in the Microsoft environment and then followed me in my course.

## 🎶 Basic work
A Query made of a formula coded in (PowerQuery) M langage is called through `ActiveSheet.ListObjects.Add` method to generate a data table from elsewhere. It is possible to get directly the data, or to manage the connection with the distant workbook before getting the data. But in this last case, documentation is especialy confusing and differences emerge between Vba functions and manual process.

The distant workbook is so described by the M formula, and have to be the same as well as its location must be correct (of course).

## 🎨 Panel
In this program, I purpose _one macro_ for each **method used**, as well as for usefull automated routine like showing or cleaning. I wrapped these auxiliary routine into one called `Management()`, in which I put in comment every routine out of interest for chosen tests.

This is how I displayed them on my fluent ribbon
<img width="1132" height="118" alt="Macro Ribbon" src="https://github.com/user-attachments/assets/11a9f03a-39bd-4d4f-9081-66b04b4a7e82" />

### Controling manualy
I considered to get **tests mixed** with _vba_ functions and the _classical_ built-in way: for this we also need to display 
- `Data` tab  >  `Queries and connections` panel

### Macro list and implementation
- **Switch to table** : not related to the other actions
  - `Create Table` : Turn the active worksheet's used range into a table set as listObject
  - `Delete Table` : Delete all present listObject in the active worksheet and recreate default columns range
 
- **exportation** :  

- **Append query** : with the M formula implemented, the generated `query` will be used within every way of importation
  
- **Import data** :
  - `Import F. External`   : One kind of importation will do by its own
  - `Import F. Connection` : Another requires _a connection_ set to process it with <ins>the creation</ins> of _a Data Model_
    
- **Set connection** :
  - `Connection Only`      ⚠ : One set a unusuable type of connection without setting Data Model
  - `Connection w. DataModel` : The other make the connection suitable with setting Data Model
    
- **Management** : This routine can be implemented after instructions in Macros setting connections or before instructions in Macros importing data, both for control or preprocessing. It wraps the sub routines below in the same order :
  - `Show  Connections`
  - `Clean Connections`
  - `Show  listObjects`
  - `Clean listObjects`

> [!NOTE]
> All names used in the previous macro are define as global variables, and so can be set or reset depending of the needs

- **Settings and variables** :
  - `Reset settings` : Reset all global variables, _for each_ one remaining empty
  - `Change settings` : Set default or access names
  - `Change columns` : Set header column's value for sheets
  - `Set Management` : Set which macros is activated in Management
  - `Execute Management` : Execute some macros in Management

> [!WARNING]
> The global variables get empty when an unhandled error occurs. For a confortable experience, it is relevant to let `Reset Setting` at the begining of each macros processing with query or exportation.

## 🧩 

