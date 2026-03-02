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
In this program, I purpose _one macro_ for each **method used**, as well as for usefull automated routine like showing or cleaning. I wrapped some of these auxiliary routine into one called `Management`, in which you can desactivate routine that would be out of interest for chosen tests. The macro managing `Queries` is more like an agent trying to get right questions to the user, out of the current list.

This is how I displayed them on my fluent ribbon
<img width="1632" height="167" alt="Macro Ribbon full" src="https://github.com/user-attachments/assets/4ff58535-a92c-4117-b738-70d883f98d98" />

### Controling manualy
I considered to get **tests mixed** with _vba_ functions and the _classical_ built-in way: for this we also need to display 
- `Data` tab  >  `Queries & connections` panel

### ***<p align="center"> or </p>***

### Macro list and implementation
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
- Manualy duplicate a query : add suffix " (2)" if suffix "(2)" is not found, else it bring the suffix to the format " (i)" with the next available index.
- Create same name query through Vba will generate an error. The macro here add +1 to the highest of first index from every names, and replace the first index of actual name. If no index is found, the new index will be 1 and added at the end of the name.
  -
- Manualy add connection : Prefix "Query - " is concatenated to query name. It should replace the potential existing one.
- Multiplicate the connection through Vba : It should add right side index starting at 1 to the name whatever it is ending by a letter or a number. Index increases to make the first available name.
  -
- 

## Special behaviors

### The Data Model log
While getting deleted with VBA and not appearing anymore in "Existing Connections" panel, _ThisWorkbookDataModel_ remains present in "Queries & Connections" panel, unlike the manual method which removes it when loading the query without adding the data to the data model or when deleting the query. Both methods will make the data model persisting as empty in the hidden connections list, and you'll have to close and re-open the workbook to not see this special connection from anywhere.

### 

