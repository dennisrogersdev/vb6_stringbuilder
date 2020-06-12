# VB6 - String Builder

### Example:

```
dim s as new StringBuilder

s.autoNewLine = True

s.Add "Select "
s.Add "   name "
s.Add "From "
s.Add "   table "

Debug.print s

s.Clear
s.Add "Teste"

Debug.print s.Text
```

