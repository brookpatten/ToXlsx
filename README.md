# ToXlsx
Extension methods to convert IList&lt;T> to Excel files using EPPlus

A Minimum Usage
```
var bytes = yourListOfT.ToXlsx()
```

With a named worksheet
```
var bytes = yourListOfT.ToWorksheet("Worksheet Name").ToXlsx();
```

Specifying Columns (previous examples columns are inferred using reflection)
```
var bytes = yourListOfT.ToWorksheet("Worksheet")
                        .WithColumn(x => x.Id, "ID")
                        .WithColumn(x => x.Name, "Name")
                        .ToXlsx();
```

Add a Title Row
```
var bytes = yourListOfT.ToWorksheet("Worksheet")
                        .WithTitle("This is the title")
                        .ToXlsx();
```

With Custom formatting
```
var bytes = yourListOfT.ToWorksheet("Title", configureHeader: f =>
                                    {
                                        f.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        f.Style.Fill.BackgroundColor.SetColor(blue);
                                        f.Style.Font.Color.SetColor(Color.White);
                                        f.Style.Font.Name = "Tahoma";
                                        f.Style.Font.Bold = true;
                                        f.Style.Font.Size = 10;
                                        f.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                        f.Style.WrapText = true;
                                    }, configureColumn: f => f.AutoFit())
                        .ToXlsx();
```

Append another worksheet
```
var bytes = yourListOfT
                .ToWorksheet("T things")
                .NextWorksheet(yourListOfK, "K things")
                .ToXlsx();
```

Custom formatting can be specified in the following ways:
* Per Worksheet you can provide formatting for header, column, header row, and per data cell
* Per Column you can provide formatting for header, column, and per data cell
