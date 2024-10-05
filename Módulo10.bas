Attribute VB_Name = "Módulo10"
Sub CREDITO()
Attribute CREDITO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CREDITO Macro
'

'
    Range("S10:S14").Select
    Selection.Copy
    Sheets("TIPO DE CAMBIO").Select
    ActiveWindow.SmallScroll Down:=440
    Range("B450").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    ActiveWindow.SmallScroll Down:=-528
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B2:B450"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=16
    Sheets("SOLICITUD CP").Select
    Range("S15").Select
End Sub
Sub ahorros()
Attribute ahorros.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ahorros Macro
'

'
    Range("Q13:Q17").Select
    Selection.Copy
    Sheets("TIPO DE CAMBIO").Select
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 167
    ActiveWindow.ScrollRow = 170
    ActiveWindow.ScrollRow = 172
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 179
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 181
    ActiveWindow.ScrollRow = 182
    ActiveWindow.ScrollRow = 184
    ActiveWindow.ScrollRow = 185
    ActiveWindow.ScrollRow = 187
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 191
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 194
    ActiveWindow.ScrollRow = 196
    ActiveWindow.ScrollRow = 197
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 200
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 204
    ActiveWindow.ScrollRow = 205
    ActiveWindow.ScrollRow = 207
    ActiveWindow.ScrollRow = 208
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 212
    ActiveWindow.ScrollRow = 213
    ActiveWindow.ScrollRow = 215
    ActiveWindow.ScrollRow = 216
    ActiveWindow.ScrollRow = 218
    ActiveWindow.ScrollRow = 220
    ActiveWindow.ScrollRow = 222
    ActiveWindow.ScrollRow = 223
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 226
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 228
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 235
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 238
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 242
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 249
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 260
    ActiveWindow.ScrollRow = 264
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 277
    ActiveWindow.ScrollRow = 286
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 304
    ActiveWindow.ScrollRow = 308
    ActiveWindow.ScrollRow = 311
    ActiveWindow.ScrollRow = 314
    ActiveWindow.ScrollRow = 317
    ActiveWindow.ScrollRow = 321
    ActiveWindow.ScrollRow = 322
    ActiveWindow.ScrollRow = 324
    ActiveWindow.ScrollRow = 327
    ActiveWindow.ScrollRow = 329
    ActiveWindow.ScrollRow = 331
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 339
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 342
    ActiveWindow.ScrollRow = 344
    ActiveWindow.ScrollRow = 346
    ActiveWindow.ScrollRow = 347
    ActiveWindow.ScrollRow = 348
    ActiveWindow.ScrollRow = 350
    ActiveWindow.ScrollRow = 351
    ActiveWindow.ScrollRow = 354
    ActiveWindow.ScrollRow = 355
    ActiveWindow.ScrollRow = 358
    ActiveWindow.ScrollRow = 360
    ActiveWindow.ScrollRow = 361
    ActiveWindow.ScrollRow = 363
    ActiveWindow.ScrollRow = 364
    ActiveWindow.ScrollRow = 365
    ActiveWindow.ScrollRow = 366
    ActiveWindow.ScrollRow = 367
    ActiveWindow.ScrollRow = 368
    ActiveWindow.ScrollRow = 369
    ActiveWindow.ScrollRow = 372
    ActiveWindow.ScrollRow = 374
    ActiveWindow.ScrollRow = 376
    ActiveWindow.ScrollRow = 380
    ActiveWindow.ScrollRow = 383
    ActiveWindow.ScrollRow = 388
    ActiveWindow.ScrollRow = 393
    ActiveWindow.ScrollRow = 399
    ActiveWindow.ScrollRow = 402
    ActiveWindow.ScrollRow = 405
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 410
    ActiveWindow.ScrollRow = 412
    ActiveWindow.ScrollRow = 413
    ActiveWindow.ScrollRow = 415
    ActiveWindow.ScrollRow = 416
    ActiveWindow.ScrollRow = 417
    ActiveWindow.ScrollRow = 419
    ActiveWindow.ScrollRow = 421
    ActiveWindow.ScrollRow = 422
    ActiveWindow.ScrollRow = 424
    ActiveWindow.ScrollRow = 425
    ActiveWindow.ScrollRow = 426
    Range("B450").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    ActiveWindow.ScrollRow = 424
    ActiveWindow.ScrollRow = 421
    ActiveWindow.ScrollRow = 419
    ActiveWindow.ScrollRow = 417
    ActiveWindow.ScrollRow = 415
    ActiveWindow.ScrollRow = 412
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 405
    ActiveWindow.ScrollRow = 399
    ActiveWindow.ScrollRow = 392
    ActiveWindow.ScrollRow = 383
    ActiveWindow.ScrollRow = 372
    ActiveWindow.ScrollRow = 359
    ActiveWindow.ScrollRow = 345
    ActiveWindow.ScrollRow = 308
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 248
    ActiveWindow.ScrollRow = 243
    ActiveWindow.ScrollRow = 241
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 239
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 223
    ActiveWindow.ScrollRow = 214
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 193
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 73
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("B2:B450"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("TIPO DE CAMBIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=8
    Sheets("CARTILLA CUENTA").Select
    Range("Q13").Select
End Sub
