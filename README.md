# denormalize
MS Excel Visual Basic Macro Utilities

Tested in MS Excel 2010

### The Problem
The "ids" column of the following table contains a list of ids separated by
commas, or line breaks. The contained macros purpose is to create a separate
row for each id in that list, i.e. the opposite of normalization, hence the
name.

|data|ids
|--- |---
|Foo |1,2
|Bar |1<br />2
|John|3
|Doe |4,5

### The Solution
- In the first step one has to selct the ids column and execute the `replaceLineBreak()` macro to replace them with comma.
- Then one has to select at least one cell in the ids column and execute the `denormalize()` macro.

The result should look like this:

|data|ids
|--- |---
|Foo |1
|Foo |1
|Bar |1
|Bar |2
|John|3
|Doe |4
|Doe |5
