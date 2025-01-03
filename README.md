## Fuzzy matching tool

GUI based tool for fuzzy string matching - currently a work in progress.

TODO:
- Implement multi column matching (only for all combinations). Current issues:
    - Loading dataset is currently weird, doesnt incorporate multi matching well.
    - Code is super messy, combine multiple functions into one maybe?
- Toggle button availabilities when running/based on other options
- Tidy up code, current things to flag:
    - Sometimes passing the dropdown to a funtion and .get()ing, sometimes passing the variable itself
    - Comments, comments and more comments.
    - Remove redudant stuff, can probably clean out packages
- Compile as an EXE