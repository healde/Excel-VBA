# Create connection to, and import data from : another workbook

## 🎼 Why ? 
It is one of the things that I wanted to add to my "how to in Vba" skills. But it appears as a less atempted option than I though and the documentation let me go forward through difficulties as the beginer as I am. So I decide to make a full snippet for looking over process and understanding the behaviors of the differents functionality around this subject.
### start
Without finding as much as solutions than I expected in forums, I recorded Macro over manual process, and get by chance more than one methods. I cleaned them out of the records context, and tried to make them work as much as I can.
### how then
After having compose a whole set of commands to run some tests, I asked for help to AI and it let me knows about some internal process in the Microsoft environment and then followed me in my course.

## 🎶 Query structure
A Query made of a formula coded in (PowerQuery) M langage is called through `ActiveSheet.ListObjects.Add` method to generate a data table from elsewhere. It is possible to get directly the data, or to manage the connection with the distant workbook before getting the data. But in this last case, documentation is especialy confusing and differences emerge between Vba functions and manual process.

The distant workbook is so described by the M formula, and have to be the same as well as its location must be correct (of course).

## 🎶 Connection type

At any moment you might want to access to the list of usable connections and data tables through a built-in dedicated panel. Here is how to display :
- `Data` _tab_  >  `Existing connections` _panel_

The process described below only use connections in the workbook. But data tables are from two types :
- Connection tables constitute the object model, and are alternative to query tables for importing with connection previously set
- Workbook tables refer directly to tables presently loaded in some cells

> [!NOTE]
> From this panel, it is possible to manually (re)load both connections or tables, but it duplicates the query source into a new one

## 🎨 Panel
In this program, I purpose _alternative macro_ to main built-in functions, plus some for usefull automated routine like showing or cleaning. I wrapped some of these auxiliary routine into one macro called `Management`, in which you can desactivate routine that would be out of interest for chosen tests. The macro managing `Queries` is more like an agent trying to get right questions to the user, out of the current list.

This is how I displayed them on my fluent ribbon
<img width="1682" height="162" alt="Macro Ribbon full" src="https://github.com/user-attachments/assets/7679a6d6-ca8f-4cb9-8b8f-862789100831" />

### Macro list and implementation

> [!WARNING]
> Before using any connection through Vba, you must once open `Data` _tab_  >  `Queries & connections` _panel_ to activate the provider ***Microsoft.Mashup.OleDB.1***

- **Switch to table** : not related to the other actions
  - `Create Table` : Turn the active worksheet's used range into a table set as listObject
  - `Delete Table` : Delete all present listObject in the active worksheet and recreate default columns range

** **

- **`Exportation`** : Make distant file from set of tables and default sheets. All is first prepared in the current workbook to be kept as control. 

- **Append or clean `Queries`** : with the M formula implemented, a generated query can be used through every kind of importation
  
- **Import data** :
  - `Import F. External`   : One kind of importation will do by its own
  - `Import F. Connection` : Another requires _a connection_ set <ins>with</ins> the creation of _a Data Model_ before process
    
- **Set connection** :
  - `Connection Only`      ⚠ : One set a unusuable type of connection without setting Data Model
  - `Connection w. DataModel` : The other make the connection suitable with setting Data Model

** **    
- **`Management`** : This routine can be implemented after instructions in Macros setting connections or before instructions in Macros importing data, both for control or preprocessing. It wraps the sub routines below in the same order :
  - `Show  Connections`
  - `Clean Connections`
  - `Show  listObjects`
  - `Clean listObjects`

- **Settings and variables** : All names used in the previous macro are define as global variables, and so can be set or reset depending of the needs :
  - `Reset settings`     : affect only empty strings
  - `Change settings`
  - `Change columns`
  - `Set Management`     : Set which macros is activated in `Management`
  - `Execute Management`

### About using global variables

> [!WARNING]
> The global variables **get empty** when an <ins>unhandled error</ins> occurs.

For a confortable experience, it is relevant to let `Reset Setting` at the begining of the macro `exportation` or any _other_ that is processing with queries. This is why it is only built to ensure the way that no empty value would be call at any moment. Whereas this code doesn't include any macro that would reset nor empty the variables what ever they are containing.

> [!NOTE]
> You can change source, target, objects default names, and columns strings through macros. But the list of sheets and number of columns in each table are currently only defined in the macro `Reset Setting`

## In-built indentation and conflicts
<br>

- _Manualy_ duplicate **sheet** : if suffix like "(i)" is found, it is filled with the next available index inside (not lowest available one), else suffix " (2)" is added.
<br>

- _Manualy_ duplicate a **query** : if suffix like "(i)" is found, it is brought to a like " (i)" format with the next available index inside (not lowest available one), else suffix " (2)" is added.
- Create same name **query** _through Vba_ method will generate an error. The macro made here add +1 to the highest number beyong all first number of each name, and replace it as new index in the actual name. If no number is found, the new index will be 1 and added at the end of the name.      
<br>

- _Manualy_ add **connection** : Prefix "Query - " is concatenated to query name. It should replace the potential existing one.
- Multiplicate the **connection** _through Vba_ : It won't replace any existing one. If the connection already exists, it adds right side to the choosen name (and without space) the lowest available index starting from 1, whatever the name is ending by a letter or a number.

### Effectless indentations

- Multiple connections can be added for a single query, but the related queryTables cannot quote a query already quoted by any other queryTable. It means that whereas a connection can refer to a same query name than an other, the following associated queryTable have to change this name if an other queryTable already refer to it. This one then adds the lowest available index to query name that it refers in a like " i" format. This case only happen within Vba, because manually adding connection from a query would replace the first connection found related to this query. But it has no effect on the loading data process, since the connection itself still links the right query.

- Importing through Vba add fantom connection _ThisWorkbookDataModel_**i** with i lowest available. This one seems to be deletable without any effect on the displayed table, as well as it is invisible from `Existing connections` panel

## Special behaviors

### Model object

Dispite remaining belong connections, the model object only display the count of connections table. This is why _ThisWorkbookDataModel_ connection is not directly deletable, neither manually nor through Vba.

> [!IMPORTANT]
> The name _ThisWorkbookDataModel_

 in `Queries & Connections` panel

** **

### Import with making queryTable that already exist will remove the old one

### cannot refresh dataTable where connections have been disabled

### Add query with space arround the name

