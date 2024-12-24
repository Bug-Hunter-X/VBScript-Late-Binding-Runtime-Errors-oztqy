Late Binding: In VBScript, if you don't explicitly declare object variables (using Dim), VBScript uses late binding. This means the type checking happens at runtime. While flexible, this can lead to runtime errors if you're using objects incorrectly or accessing non-existent properties.  Example:

```vbscript
Set obj = CreateObject("Some.Object")
' No Dim statement, late binding is used
result = obj.NonExistentMethod()
```
This might only throw an error when NonExistentMethod() is called during execution, not during compile time. 