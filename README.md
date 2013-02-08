Classic ASP Mustache
====================

A template engine for classic ASP that implements basic Mustache syntax.

Example
-------

```asp
Set M = new Mustache

Set dict=Server.CreateObject("Scripting.Dictionary")
dim output, template

template = _
    "Hello {{name}} " & _
    "You have just won ${{value}}! " & _
    "{{#in_ca}} " & _
    "Well, ${{taxed_value}}, after taxes. "  &_
    "{{/in_ca}}"

dict.Add "name", "Chris"
dict.Add "value", 10000
dict.Add "taxed_value", 10000 - (10000 * 0.4)
dict.Add "in_ca", true

output = M.render(template, dict)
Response.Write output
```

License
-------
MIT