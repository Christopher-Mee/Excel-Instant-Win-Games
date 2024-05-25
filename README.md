# Excel Instant Win Games

An Excel sheet used to track and enter Instant Win Games (IWG).

### Features ###

* VBA Macros
  * When opening the direct link to a game, a macro will automatically add a check mark.
    * This helps indicate which games you have already entered today.
    * The check mark is made using wingdings font type.
      * 'P' = check mark (success)
      * 'O' = x mark (error) 
* Conditional Formatting
  * The cells for the start and end date are color coded to indicate if a contest is active.
    * If the cells are green, the contest has started.
    * If the cells are red, the contest has ended.
    * If the cells are orange, the contest has yet to start.
  * Rows with data have a border applied.
* Hyperlinks
  * The Color style has been modified to keep hyperlinks from changing color when opened.
 
### Tips ###

* To add or edit a hyperlink in a cell
  * Use 'ctrl' + 'k'
  * If Excel freezes when adding a hyperlink, close you browser. (Bug with Google Chrome?)
* Once the Excel Sheet has been filled up
  * It is tedious to scroll down and add a new entry. 
  * Use 'shift' + 'space', then 'ctrl' + 'numpad +' to add a new entry at the top of the Sheet.
  * Alternatively use 'shift' + 'end' to go to the end of all entries.
  * Then use 'shift' + 'home' to go to the top of all entries.
* To sort the IWG list
  * Custom sorting will help to properly organize your games based on any column you wish.
  * Use 'context menu key' + 'o' + 'u' to open the custom sorting window.
