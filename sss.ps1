$word = New-Object -ComObject Word.Application
$word.Visible = $True
$doc = $word.Documents.Add()
$Selection = $word.Selection
$Selection.TypeParagraph()
$Selection.Font.Name = "Times New Roman"
$Selection.Font.Size = 18
$Selection.TypeText("Hello world!")
