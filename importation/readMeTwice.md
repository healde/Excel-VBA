# <p align="center">SandBox</p>
# Figure out connections and basic data transferts between Excel files

## 🎼 Why ? 
It is one of the things that I wanted to add to my "how to in Vba" skills. But it appears less as a mild option than I though and the documentation was difficult for me to get through, as the beginer as I am. So I decide to make a program for looking over general process and understanding the behaviors of the differents functionality around this subject.
### starting line
Without finding as much as solutions than I expected in forums, I recorded Macro ⭕ over manual process, and get by chance more than one method. I cleaned them out of the records context 🧹 and tried to make them work as much as I can.
### how then
After having composed a whole set of commands to run some tests, I also asked for help to AI and it let me know about some _internal process_ in the Microsoft environment. It has then followed me in my course, while I pursuid some running tests and the development of **control process**.

## 🎶 So
Through this ReadMe and with this program share, I try to offer a collection of my observations and an attempt of supplement to the documentation.

<br>
<br>
<br>

## Excel Basis

### Queries and providers

In the inital phaze, I builded the code over only one provider. it is near the end that I looked for other options, and this program present now <ins>two providers</ins> which offer similar capabilities for fetching data through one <ins>excel file</ins> to an other. <br>
<br>
A provider is in charge of managing a _query_ and **ensure the _connection_** to a source for importing the right data. 

> [!NOTE]
> **Connections and Data Tables** : Getting any imported data into a displayed table is a complet other step after processing a connection, even main functionnalities use two in one process.

The resulting `.workbookConnection` object which is accessible within Vba also get `.connectionType` property. Each provider get a **list of source type** which are possible to connect with.
<br>
<br>
Source : [Vba XlConnectionType enumeration](https://learn.microsoft.com/en-gb/office/vba/api/excel.xlconnectiontype)
| Name | Value | Description |
| :--- | :---: | --- |
| xlConnectionTypeOLEDB | 1 | OLEDB |
| xlConnectionTypeODBC | 2 | ODBC |
| xlConnectionTypeXMLMAP | 3 | XML MAP |
| xlConnectionTypeTEXT | 4 | Text |
| xlConnectionTypeWEB | 5 | Web |
| xlConnectionTypeDATAFEED | 6 | Data Feed |
| xlConnectionTypeMODEL | 7 | PowerPivot Model |
| xlConnectionTypeWORKSHEET | 8 | Worksheet |
| xlConnectionTypeNOSOURCE | 9 | No source |

ACE and MashUp create connection of type 1 corresponding to OleDB



## 🎨 Panel

**In this program**, I purpose to try _alternative_ macro to main _built-in functions_, plus some for _usefull automated routine_ like showing or cleaning. I wrapped some of these auxiliary routines into one macro called `Management`, in which you can desactivate routines that would be out of interest for chosen tests. Indeed this one is then called as a preprocess routine as well.

The macro managing `Queries` is more like an agent trying to get right questions to the user, out of the current list. There are also macros to access some variables and change functions target and settings.

<br>

This is how I displayed them on my fluent ribbon

<img width="1630" height="167" alt="Macro Ribbon full" src="https://github.com/user-attachments/assets/931afbc1-4711-4e6c-bc88-cc9a119f4934" />

### 🖼 Macro list and implementation

> [!IMPORTANT]
> Before using any connection through Vba, you must activate the provider ***Microsoft.Mashup.OleDB.1*** by opening once :
> - `Data` _tab_  >  `Queries & connections` _panel_

- **Switch to table** (not related to the other actions) :
  - `Create Table` : Turn the active worksheet's used range into a table set as listObject
  - `Delete Table` : Delete all present listObject in the active worksheet and recreate default columns range

** **

- **`Exportation`** : Make distant file from set of tables and default sheets. All is first prepared in the current workbook to be kept as control. 

- **Append or clean `Queries`** : with the M formula implemented, a generated query can be used through every kind of importation
  
- **Import data** :
  - `Import F. External`   : One kind of importation will do by its own
  - `Import F. Connection` : Another requires _a connection_ <ins>refered</ins> in _Data Model_ object collection before processing
    
- **Set connection** :
  - `Connection w. Option`    : Set first the connection only. To load its reference into Data Model is in option
  - `Connection w. DataModel` : Make directly the new connection suitable to import by loading its reference into Data Model

** **    
- **`Auto Management`** : This routine can be implemented after instructions in Macros setting connections or before instructions in Macros importing data, both for control or preprocessing. It wraps the sub routines below in the same order :
  - `Show  Connections`
  - `Clean Connections`
  - `Show  listObjects`
  - `Clean listObjects`

- **Settings and variables** : All names used in the previous macro are define as global variables, and so can be set or reset depending of the needs :
  - `Reset settings`     : affect only empty strings
  - `Change settings`
  - `Change columns`
  - `Set Management`     : Set which macros is activated in `Auto Management`
  - `Get Management`

### 🎚 About using global variables

> [!WARNING]
> The global variables **get empty** when an <ins>unhandled error</ins> occurs.

For a confortable experience, it is relevant to let `Reset Setting` at the begining of the macro `exportation` or any _other_ that is processing with queries. This is why it is only built to ensure the way that no empty value would be call at any moment. Whereas this code doesn't include any macro that would reset nor empty the variables what ever they are containing.

> [!TIP]
> You can change source, target, objects default names, and columns strings through macros. But the list of sheets and number of columns in each table are currently only defined in the macro `Reset Setting` in <ins>ThisWorkbook.cls<ins>.

<br>
<br>

## ➕ In-built indentation and conflicts

- _Manualy_ duplicate **sheet** : if suffix like "(i)" is found, it is filled with the next available index inside (not lowest available one), else suffix " (2)" is added[^1].
<br>

- _Manualy_ duplicate a **query** : if suffix like "(i)" is found, it is brought to a like " (i)" format with the next available index inside (not lowest available one), else suffix " (2)" is added[^1].
- Create same name **query** _through Vba_ method will generate an error. The macro made here add +1 to the highest number among all first number of each name, and replace it as new index in the actual name. If no number is found, the new index will be 1 and added at the end of the name.      
<br>

- _Manualy_ add **connection** : Prefix "Query - " is concatenated to query name. It should replace the potential existing one.
<a name="connection-through-vba"></a>
- Multiplicate the **connection** _through Vba_ : It won't replace any existing one. If the connection already exists, it adds right side to the choosen name (and without space) the lowest available index starting from 1, whatever the name is ending by a letter or a number. The `.name` parameter can be set to empty string (""), then the default name is "Connection". 
<br>

- Repeat **importation** _through Vba_ : Names are not automaticaly indented, if they are define in `.workbookConnection` and `.displayedName` properties. Else table's name is always "Table_ExternalData_**i**", with i is initiated to 1 and follows the actual highest indentation value. When `.SourceType` is set to _xlSrcModel_, the connection's name follows the previous format and is set like "ModelConnection_ExternalData_**i**", while i comes actually from the indentation value of its related data table. However when `.SourceType` is set to _xlSrcExternal_, the new connection that is created follows the same rule than simple connections for their names, [see above](#connection-through-vba).

### Effectless indentations

- Importing from connection through Vba adds fantom connection _ThisWorkbookDataModel_**i** with i lowest available. This one seems to be deletable without any effect on the displayed table, as well as it is invisible from the `Existing connections` panel.

- Multiple connections can be added for a single query within Vba. Every table have to get its own name as single identifier, so connection tables cannot share the same name even their related connections are reaching the same query. This table name is by default the same as the query name. the lowest available index in a like " i" format is added to this default name to set it as a single name.
> [!NOTE]
> This case only happen within Vba, because adding a connection manually does not multiplicate connections to one query (see [below](#add-manually)). But it has no effect on the loading data process, since the connection itself still links the right query.

## 🖌 Special behaviors

### Add query with spaces arround the name

When you rename a query name or when you use it through any process refering to it, start and end spaces should be then escaped. But when you add query within Vba, so you can get multiple same looking queries (which are not). However you should then not be able to rename a query like another nor to confirm a name with spaces arround, and you can refer to a query only if it exists with a conform name.

### Forbidden caracters in table name

Workbook tables that would have been created through a connecting process get auto-generated name. If this one present any forbidden caracters such as space or parenthesis, they are replaced by underscore. Unless the name ends by forbidden caracters, these are then removed.

### Manually add connection, versus doing within VBA

Connection is set _by hand_ with default name and default description. Both come from the name of the query which the connection support. Name and description are arguments of the connections _`.add` method_, so you can choose to let a different sentence for the description as a hint of _origin_.

<a name="add-manually"></a>
> [!IMPORTANT]
> For _hand-commands_ is always treated the **first** matching **connection of the list**.

- It means to add connection manually from a query would creates new connection, only if none of connections already refers to this query. Else it updates the options of the first connection found, such as setting a connection table into Data Model, while it refresh with preserving original name and description.

- As well, renaming query manually will only affect the first connection which refers to it. This action will not snap the link to its related connection table if this one exists in the model object. _However_, unlike setting connection again, the affected connection seems to at least partially reset. Indeed name and description both change for the default ones availables, and only the connection table's name would stay unchanged[^2].

- However, importing data through connection within Vba seems to finalise renaming query affectations by updating the connection table's name as well. This behavior is far than updating while it not concerns indented names when lower default name become available again[^3].

> [!WARNING]
> Delete a query manually erase all connections and their tables bound to Data Model, while the connections entirely remain if a query is deleted through Vba.

### Data tables are unrefreshable if some of their connections have been disabled
- Data table imported from an existing connection needs this connection as much as its internal workbook connection made for it.
- Whereas data table imported directly from an external source needs so far its workbook connection only.

# Bank of pictures

[^1]:<img alt="C - Duplicate Queries" src="https://github.com/user-attachments/assets/7cb3992b-0266-45a6-8dc3-45407f82c618" />

[^2]:<img width="80%" height="80%" alt="Succesvely load to and unload from  Data Model" src="https://github.com/user-attachments/assets/1d3fd47b-6329-4fdc-a6ec-60a656f24e68" />

[^3]:<img width="80%" height="80%" alt="Rename query then import" src="https://github.com/user-attachments/assets/6d2341be-b511-4c03-9507-d492d6e591e1" />
