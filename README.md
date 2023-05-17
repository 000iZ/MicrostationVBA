## Automating Microstation with VBA
A VBA script to automate the editing of elements in **Bentley Microstation CE**

This script will change the color and line weights of _all_ elements on a level you specify. Should any errors occurr, detailed error messages will be displayed in Microstation itself.

If used with Batch Process, multiple files can be modified automatically. Running time for each file is around one-tenth of a second.a

CAD best practices with regards to level settings will also be updated, where no elements will be allowed to use custom settings separate to its level's default. Any further changes to the level's settings will _populate through all elements_ on that level.

## Remarks
Additional information on use can be found within the script itself.

Completed and implemented at the Toronto Transit Commision in early 2023.
