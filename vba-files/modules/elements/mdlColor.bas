Attribute VB_Name = "mdlColor"
' Public Theme, ThemeIcon
' Public NavFrameColor        As Long
' Public NavMenuColor         As Long
' Public NavForeColor         As Long
' Public RowBackBackColor     As Long
' Public HeaderColor          As Long
' Public topBarInfo           As Long

' Public ActiveBackLeftMenu   As Long
' Public MenuLabel            As Long
' Public MenuBorder           As Long
' Public BackLabelMenu        As Long

' Public SortIconHead         As Long
' Public SortIconHover        As Long
' Public RowBackground        As Long
' Public EditIconActive       As Long
' Public EditIconHover        As Long
' Public DeleteIconHover      As Long
' Public DeleteIconActive     As Long
' Public ActionLabelBack      As Long
' Public ActivePagination     As Long
' Public PaginationHover      As Long

' Function Colors() As Long
'     Dim r As Integer, g As Integer, b As Integer
'     r = Int(Rnd * 256)
'     g = Int(Rnd * 256)
'     b = Int(Rnd * 256)
'     Colors = RGB(r, g, b)
' End Function

' Sub MenuThemeColor(themeType As String) 
'     'Theme = themeType in Sub MenuThemeColor
'     If themeType = "Dark" Then
'         NavFrameColor           = GenHexToRGB("#f0f0f0") 'RGB(240,240,240) '#f0f0f0
'         NavForeColor            = GenHexToRGB("#C8C8C8") 'RGB(200,200,200) '#C8C8C8
'         topBarInfo              = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         RowBackBackColor        = GenHexToRGB("#f0f0f0") 'RGB(240,240,240) '#f0f0f0
'         HeaderColor             = GenHexToRGB("#2e2e2e") 'RGB(46,46,46) '#2e2e2e
'         ActivePagination        = GenHexToRGB("#a9aeb1") 'RGB(169, 174, 177)  '#a9aeb1
'         PaginationHover         = GenHexToRGB("#71767a") 'RGB(113, 118, 122)  '#71767a

'         ActiveBackLeftMenu      = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#191b1b
'         MenuLabel               = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         MenuBorder              = GenHexToRGB("#c9c9c9") 'RGB(201,201,201) '#c9c9c9
'         BackLabelMenu           = GenHexToRGB("#e6e6e6") 'RGB(230,230,230) '#e6e6e6

'         SortIconHead            = GenHexToRGB("#98afd2") 'RGB(152,175,210) '#98afd2
'         SortIconHover           = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         RowBackground           = GenHexToRGB("#f6f6f6") 'RGB(246,246,246) '#f6f6f6
'         EditIconActive          = GenHexToRGB("#98afd2") 'RGB(152,175,210) '#98afd2
'         EditIconHover           = GenHexToRGB("#345d96") 'RGB(52,93,150) '	#345d96
'         DeleteIconHover         = GenHexToRGB("#fb5a47") 'RGB(251,90,71) '#fb5a47
'         DeleteIconActive        = GenHexToRGB("#f7bbb1") 'RGB(247,187,177) '#f7bbb1
'         ActionLabelBack         = GenHexToRGB("#1a4480") 'RGB(26,68,128) '#1a4480
'     ElseIf themeType = "Light" Then
'         NavFrameColor           = vbWhite
'         topBarInfo              = vbWhite
'         NavForeColor            = GenHexToRGB("#1b2b85")  'RGB(27,43,133) '#1b2b85
'         RowBackBackColor        = GenHexToRGB("#f0f0f0")  'RGB(240,240,240) '#f0f0f0
'         HeaderColor             = GenHexToRGB("#2e2e2e")  'RGB(46,46,46) '#2e2e2e
'         ActivePagination        = GenHexToRGB("#a9aeb1") 'RGB(169, 174, 177)  '#a9aeb1
'         PaginationHover         = GenHexToRGB("#71767a") 'RGB(113, 118, 122)  '#71767a
        
'         ActiveBackLeftMenu      = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         MenuLabel               = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         MenuBorder              = GenHexToRGB("#c9c9c9") 'RGB(201,201,201) '#c9c9c9
'         BackLabelMenu           = GenHexToRGB("#e6e6e6") 'RGB(230,230,230) '#e6e6e6

'         SortIconHead            = GenHexToRGB("#98afd2") 'RGB(152,175,210) '#98afd2
'         SortIconHover           = GenHexToRGB("#005ea2") 'RGB(0,94,162) '#005ea2
'         RowBackground           = GenHexToRGB("#f6f6f6") 'RGB(246,246,246) '#f6f6f6
'         EditIconActive          = GenHexToRGB("#98afd2") 'RGB(152,175,210) '#98afd2
'         EditIconHover           = GenHexToRGB("#345d96") 'RGB(52,93,150) '	#345d96
'         DeleteIconHover         = GenHexToRGB("#fb5a47") 'RGB(251,90,71) '#fb5a47
'         DeleteIconActive        = GenHexToRGB("#f7bbb1") 'RGB(247,187,177) '#f7bbb1
'         ActionLabelBack         = GenHexToRGB("#1a4480") 'RGB(26,68,128) '#1a4480
'     End If
'     Theme = themeType
' End Sub

' Public Function HexToRGB(hexColor As String) As String
'     ' Purpose: Convert a Hex color code (e.g., "#FFAA00" or "FFAA00") to a VBA RGB Long value
'     Dim R As Integer
'     Dim G As Integer
'     Dim B As Integer
'     Dim hexValue As String
    
'     ' Remove the "#" if it exists
'     hexValue = Replace(UCase(hexColor), "#", "")
    
'     ' Ensure the string is 6 characters long (pad with leading zeros if necessary)
'     hexValue = Right$("000000" & hexValue, 6)
    
'     ' Extract the R, G, and B components and convert them from hex to decimal
'     ' Val("&H" & ...) interprets the hex string as a hexadecimal number
'     R = Val("&H" & Mid(hexValue, 1, 2))
'     G = Val("&H" & Mid(hexValue, 3, 2))
'     B = Val("&H" & Mid(hexValue, 5, 2))
    
'     ' Use the built-in RGB function to return the final VBA color value (Long integer)
'     HexToRGB = "RGB(" & R & "," &  G & "," & B & ")" 
' End Function

' Public Sub SetColorRGB()
'     Dim ws As Worksheet
'     Dim colorRange As Range
'     Dim i, useRangeId As Integer
'     Dim Col As Range

'     Set ws = Sheets("Color")
'     Set colorRange = ActiveSheet.Range("A3:F462")
'     useRangeId = ws.Cells(Rows.Count, "D").End(xlUp).Row - 2
'     For i = 1 To useRangeId
'         For Each Col In colorRange.Columns
'             If (Col.Cells(1, 1).Address = "$D$3") Then 
'                 ws.Range("C" & i+2).Interior.Color = GenHexToRGB(Col.Cells(i, 1).Value)
'             End If
'         Next Col
'     Next i
' End Sub

' Public Function GenHexToRGB(hexColor As String) As Long
'     Dim R As Integer
'     Dim G As Integer
'     Dim B As Integer
'     Dim hexValue As String
    
'     ' Remove the "#" if it exists
'     hexValue = Replace(UCase(hexColor), "#", "")
    
'     ' Ensure the string is 6 characters long (pad with leading zeros if necessary)
'     hexValue = right$("000000" & hexValue, 6)
    
'     ' Extract the R, G, and B components and convert them from hex to decimal
'     ' Val("&H" & ...) interprets the hex string as a hexadecimal number
'     R = Val("&H" & Mid(hexValue, 1, 2))
'     G = Val("&H" & Mid(hexValue, 3, 2))
'     B = Val("&H" & Mid(hexValue, 5, 2))
    
'     ' Use the built-in RGB function to return the final VBA color value (Long integer)
'     GenHexToRGB = RGB(R, G, B)
' End Function

' No	ColorType	    ColorHex	 RGB
' ******************************************************
' 1	    Red cool		#f8eff1	RGB(248,239,241)
' 2	    Red cool		#f3e1e4	RGB(243,225,228)
' 3	    Red cool		#ecbec6	RGB(236,190,198)
' 4	    Red cool		#e09aa6	RGB(224,154,166)
' 5	    Red cool		#e16b80	RGB(225,107,128)
' 6	    Red cool		#cd425b	RGB(205,66,91)
' 7	    Red cool		#9e394b	RGB(158,57,75)
' 8	    Red cool		#68363f	RGB(104,54,63)
' 9	    Red cool		#40282c	RGB(64,40,44)
' 10	Red cool		#1e1517	RGB(30,21,23)
' 11	Red cool vivid		#fff2f5	RGB(255,242,245)
' 12	Red cool vivid		#f8dfe2	RGB(248,223,226)
' 13	Red cool vivid		#f8b9c5	RGB(248,185,197)
' 14	Red cool vivid		#fd8ba0	RGB(253,139,160)
' 15	Red cool vivid		#f45d79	RGB(244,93,121)
' 16	Red cool vivid		#e41d3d	RGB(228,29,61)
' 17	Red cool vivid		#b21d38	RGB(178,29,56)
' 18	Red cool vivid		#822133	RGB(130,33,51)
' 19	Red cool vivid		#4f1c24	RGB(79,28,36)
' 20	Red		#f9eeee	RGB(249,238,238)
' 21	Red		#f8e1de	RGB(248,225,222)
' 22	Red		#f7bbb1	RGB(247,187,177)
' 23	Red		#f2938c	RGB(242,147,140)
' 24	Red		#e9695f	RGB(233,105,95)
' 25	Red		#d83933	RGB(216,57,51)
' 26	Red		#a23737	RGB(162,55,55)
' 27	Red		#6f3331	RGB(111,51,49)
' 28	Red		#3e2927	RGB(62,41,39)
' 29	Red		#1b1616	RGB(27,22,22)
' 30	Red warm		#f6efea	RGB(246,239,234)
' 31	Red warm		#f4e3db	RGB(244,227,219)
' 32	Red warm		#ecc0a7	RGB(236,192,167)
' 33	Red warm		#dca081	RGB(220,160,129)
' 34	Red warm		#d27a56	RGB(210,122,86)
' 35	Red warm		#c3512c	RGB(195,81,44)
' 36	Red warm		#805039	RGB(128,80,57)
' 37	Red warm		#524236	RGB(82,66,54)
' 38	Red warm		#332d29	RGB(51,45,41)
' 39	Red warm		#1f1c18	RGB(31,28,24)
' 40	Red warm vivid		#fff5ee	RGB(255,245,238)
' 41	Red warm vivid		#fce1d4	RGB(252,225,212)
' 42	Red warm vivid		#f6bd9c	RGB(246,189,156)
' 43	Red warm vivid		#f39268	RGB(243,146,104)
' 44	Red warm vivid		#ef5e25	RGB(239,94,37)
' 45	Red warm vivid		#d54309	RGB(213,67,9)
' 46	Red warm vivid		#9c3d10	RGB(156,61,16)
' 47	Red warm vivid		#63340f	RGB(99,52,15)
' 48	Red warm vivid		#3e2a1e	RGB(62,42,30)
' 49	Red vivid		#fff3f2	RGB(255,243,242)
' 50	Red vivid		#fde0db	RGB(253,224,219)
' 51	Red vivid		#fdb8ae	RGB(253,184,174)
' 52	Red vivid		#ff8d7b	RGB(255,141,123)
' 53	Red vivid		#fb5a47	RGB(251,90,71)
' 54	Red vivid		#e52207	RGB(229,34,7)
' 55	Red vivid		#b50909	RGB(181,9,9)
' 56	Red vivid		#8b0a03	RGB(139,10,3)
' 57	Red vivid		#5c1111	RGB(92,17,17)
' 58	Orange warm		#faeee5	RGB(250,238,229)
' 59	Orange warm		#fbe0d0	RGB(251,224,208)
' 60	Orange warm		#f7bca2	RGB(247,188,162)
' 61	Orange warm		#f3966d	RGB(243,150,109)
' 62	Orange warm		#e17141	RGB(225,113,65)
' 63	Orange warm		#bd5727	RGB(189,87,39)
' 64	Orange warm		#914734	RGB(145,71,52)
' 65	Orange warm		#633a32	RGB(99,58,50)
' 66	Orange warm		#3d2925	RGB(61,41,37)
' 67	Orange warm		#1c1615	RGB(28,22,21)
' 68	Orange warm vivid		#fff3ea	RGB(255,243,234)
' 69	Orange warm vivid		#ffe2d1	RGB(255,226,209)
' 70	Orange warm vivid		#fbbaa7	RGB(251,186,167)
' 71	Orange warm vivid		#fc906d	RGB(252,144,109)
' 72	Orange warm vivid		#ff580a	RGB(255,88,10)
' 73	Orange warm vivid		#cf4900	RGB(207,73,0)
' 74	Orange warm vivid		#a72f10	RGB(167,47,16)
' 75	Orange warm vivid		#782312	RGB(120,35,18)
' 76	Orange warm vivid		#3d231d	RGB(61,35,29)
' 77	Orange		#f6efe9	RGB(246,239,233)
' 78	Orange		#f2e4d4	RGB(242,228,212)
' 79	Orange		#f3bf90	RGB(243,191,144)
' 80	Orange		#f09860	RGB(240,152,96)
' 81	Orange		#dd7533	RGB(221,117,51)
' 82	Orange		#a86437	RGB(168,100,55)
' 83	Orange		#775540	RGB(119,85,64)
' 84	Orange		#524236	RGB(82,66,54)
' 85	Orange		#332d27	RGB(51,45,39)
' 86	Orange		#1b1614	RGB(27,22,20)
' 87	Orange vivid		#fef2e4	RGB(254,242,228)
' 88	Orange vivid		#fce2c5	RGB(252,226,197)
' 89	Orange vivid		#ffbc78	RGB(255,188,120)
' 90	Orange vivid		#fa9441	RGB(250,148,65)
' 91	Orange vivid		#e66f0e	RGB(230,111,14)
' 92	Orange vivid		#c05600	RGB(192,86,0)
' 93	Orange vivid		#8c471c	RGB(140,71,28)
' 94	Orange vivid		#5f3617	RGB(95,54,23)
' 95	Orange vivid		#352313	RGB(53,35,19)
' 96	Gold		#f5f0e6	RGB(245,240,230)
' 97	Gold		#f1e5cd	RGB(241,229,205)
' 98	Gold		#dec69a	RGB(222,198,154)
' 99	Gold		#c7a97b	RGB(199,169,123)
' 100	Gold		#ad8b65	RGB(173,139,101)
' 101	Gold		#8e704f	RGB(142,112,79)
' 102	Gold		#6b5947	RGB(107,89,71)
' 103	Gold		#4d4438	RGB(77,68,56)
' 104	Gold		#322d26	RGB(50,45,38)
' 105	Gold		#191714	RGB(25,23,20)
' 106	Gold vivid		#fef0c8	RGB(254,240,200)
' 107	Gold vivid		#ffe396	RGB(255,227,150)
' 108	Gold vivid		#ffbe2e	RGB(255,190,46)
' 109	Gold vivid		#e5a000	RGB(229,160,0)
' 110	Gold vivid		#c2850c	RGB(194,133,12)
' 111	Gold vivid		#936f38	RGB(147,111,56)
' 112	Gold vivid		#7a591a	RGB(122,89,26)
' 113	Gold vivid		#5c410a	RGB(92,65,10)
' 114	Gold vivid		#3b2b15	RGB(59,43,21)
' 115	Yellow		#faf3d1	RGB(250,243,209)
' 116	Yellow		#f5e6af	RGB(245,230,175)
' 117	Yellow		#e6c74c	RGB(230,199,76)
' 118	Yellow		#c9ab48	RGB(201,171,72)
' 119	Yellow		#a88f48	RGB(168,143,72)
' 120	Yellow		#8a7237	RGB(138,114,55)
' 121	Yellow		#6b5a39	RGB(107,90,57)
' 122	Yellow		#504332	RGB(80,67,50)
' 123	Yellow		#332d27	RGB(51,45,39)
' 124	Yellow		#1a1614	RGB(26,22,20)
' 125	Yellow vivid		#fff5c2	RGB(255,245,194)
' 126	Yellow vivid		#fee685	RGB(254,230,133)
' 127	Yellow vivid		#face00	RGB(250,206,0)
' 128	Yellow vivid		#ddaa01	RGB(221,170,1)
' 129	Yellow vivid		#b38c00	RGB(179,140,0)
' 130	Yellow vivid		#947100	RGB(148,113,0)
' 131	Yellow vivid		#776017	RGB(119,96,23)
' 132	Yellow vivid		#5c4809	RGB(92,72,9)
' 133	Yellow vivid		#422d19	RGB(66,45,25)
' 134	Green warm		#f1f4d7	RGB(241,244,215)
' 135	Green warm		#e7eab7	RGB(231,234,183)
' 136	Green warm		#cbd17a	RGB(203,209,122)
' 137	Green warm		#a6b557	RGB(166,181,87)
' 138	Green warm		#8a984b	RGB(138,152,75)
' 139	Green warm		#6f7a41	RGB(111,122,65)
' 140	Green warm		#5a5f38	RGB(90,95,56)
' 141	Green warm		#45472f	RGB(69,71,47)
' 142	Green warm		#2d2f21	RGB(45,47,33)
' 143	Green warm		#171712	RGB(23,23,18)
' 144	Green warm vivid		#f5fbc1	RGB(245,251,193)
' 145	Green warm vivid		#e7f434	RGB(231,244,52)
' 146	Green warm vivid		#c5d30a	RGB(197,211,10)
' 147	Green warm vivid		#a3b72c	RGB(163,183,44)
' 148	Green warm vivid		#7e9c1d	RGB(126,156,29)
' 149	Green warm vivid		#6a7d00	RGB(106,125,0)
' 150	Green warm vivid		#5a6613	RGB(90,102,19)
' 151	Green warm vivid		#4b4e10	RGB(75,78,16)
' 152	Green warm vivid		#38380b	RGB(56,56,11)
' 153	Green		#eaf4dd	RGB(234,244,221)
' 154	Green		#dfeacd	RGB(223,234,205)
' 155	Green		#b8d293	RGB(184,210,147)
' 156	Green		#9bb672	RGB(155,182,114)
' 157	Green		#7d9b4e	RGB(125,155,78)
' 158	Green		#607f35	RGB(96,127,53)
' 159	Green		#4c6424	RGB(76,100,36)
' 160	Green		#3c4a29	RGB(60,74,41)
' 161	Green		#293021	RGB(41,48,33)
' 162	Green		#161814	RGB(22,24,20)
' 163	Green vivid		#ddf9c7	RGB(221,249,199)
' 164	Green vivid		#c5ee93	RGB(197,238,147)
' 165	Green vivid		#98d035	RGB(152,208,53)
' 166	Green vivid		#7fb135	RGB(127,177,53)
' 167	Green vivid		#719f2a	RGB(113,159,42)
' 168	Green vivid		#538200	RGB(83,130,0)
' 169	Green vivid		#466c04	RGB(70,108,4)
' 170	Green vivid		#2f4a0b	RGB(47,74,11)
' 171	Green vivid		#243413	RGB(36,52,19)
' 172	Green cool		#ecf3ec	RGB(236,243,236)
' 173	Green cool		#dbebde	RGB(219,235,222)
' 174	Green cool		#b4d0b9	RGB(180,208,185)
' 175	Green cool		#86b98e	RGB(134,185,142)
' 176	Green cool		#5e9f69	RGB(94,159,105)
' 177	Green cool		#4d8055	RGB(77,128,85)
' 178	Green cool		#446443	RGB(68,100,67)
' 179	Green cool		#37493b	RGB(55,73,59)
' 180	Green cool		#28312a	RGB(40,49,42)
' 181	Green cool		#1a1f1a	RGB(26,31,26)
' 182	Green cool vivid		#e3f5e1	RGB(227,245,225)
' 183	Green cool vivid		#b7f5bd	RGB(183,245,189)
' 184	Green cool vivid		#70e17b	RGB(112,225,123)
' 185	Green cool vivid		#21c834	RGB(33,200,52)
' 186	Green cool vivid		#00a91c	RGB(0,169,28)
' 187	Green cool vivid		#008817	RGB(0,136,23)
' 188	Green cool vivid		#216e1f	RGB(33,110,31)
' 189	Green cool vivid		#154c21	RGB(21,76,33)
' 190	Green cool vivid		#19311e	RGB(25,49,30)
' 191	Mint		#dbf6ed	RGB(219,246,237)
' 192	Mint		#c7efe2	RGB(199,239,226)
' 193	Mint		#92d9bb	RGB(146,217,187)
' 194	Mint		#5abf95	RGB(90,191,149)
' 195	Mint		#34a37e	RGB(52,163,126)
' 196	Mint		#2e8367	RGB(46,131,103)
' 197	Mint		#286846	RGB(40,104,70)
' 198	Mint		#204e34	RGB(32,78,52)
' 199	Mint		#193324	RGB(25,51,36)
' 200	Mint		#0d1a12	RGB(13,26,18)
' 201	Mint vivid		#c9fbeb	RGB(201,251,235)
' 202	Mint vivid		#83fcd4	RGB(131,252,212)
' 203	Mint vivid		#0ceda6	RGB(12,237,166)
' 204	Mint vivid		#04c585	RGB(4,197,133)
' 205	Mint vivid		#00a871	RGB(0,168,113)
' 206	Mint vivid		#008659	RGB(0,134,89)
' 207	Mint vivid		#146947	RGB(20,105,71)
' 208	Mint vivid		#0c4e29	RGB(12,78,41)
' 209	Mint vivid		#0d351e	RGB(13,53,30)
' 210	Mint cool		#e0f7f6	RGB(224,247,246)
' 211	Mint cool		#c4eeeb	RGB(196,238,235)
' 212	Mint cool		#9bd4cf	RGB(155,212,207)
' 213	Mint cool		#6fbab3	RGB(111,186,179)
' 214	Mint cool		#4f9e99	RGB(79,158,153)
' 215	Mint cool		#40807e	RGB(64,128,126)
' 216	Mint cool		#376462	RGB(55,100,98)
' 217	Mint cool		#2a4b45	RGB(42,75,69)
' 218	Mint cool		#203131	RGB(32,49,49)
' 219	Mint cool		#111818	RGB(17,24,24)
' 220	Mint cool vivid		#d5fbf3	RGB(213,251,243)
' 221	Mint cool vivid		#7efbe1	RGB(126,251,225)
' 222	Mint cool vivid		#29e1cb	RGB(41,225,203)
' 223	Mint cool vivid		#1dc2ae	RGB(29,194,174)
' 224	Mint cool vivid		#00a398	RGB(0,163,152)
' 225	Mint cool vivid		#008480	RGB(0,132,128)
' 226	Mint cool vivid		#0f6460	RGB(15,100,96)
' 227	Mint cool vivid		#0b4b3f	RGB(11,75,63)
' 228	Mint cool vivid		#123131	RGB(18,49,49)
' 229	Cyan		#e7f6f8	RGB(231,246,248)
' 230	Cyan		#ccecf2	RGB(204,236,242)
' 231	Cyan		#99deea	RGB(153,222,234)
' 232	Cyan		#5dc0d1	RGB(93,192,209)
' 233	Cyan		#449dac	RGB(68,157,172)
' 234	Cyan		#168092	RGB(22,128,146)
' 235	Cyan		#2a646d	RGB(42,100,109)
' 236	Cyan		#2c4a4e	RGB(44,74,78)
' 237	Cyan		#203133	RGB(32,49,51)
' 238	Cyan		#111819	RGB(17,24,25)
' 239	Cyan vivid		#e5faff	RGB(229,250,255)
' 240	Cyan vivid		#a8f2ff	RGB(168,242,255)
' 241	Cyan vivid		#52daf2	RGB(82,218,242)
' 242	Cyan vivid		#00bde3	RGB(0,189,227)
' 243	Cyan vivid		#009ec1	RGB(0,158,193)
' 244	Cyan vivid		#0081a1	RGB(0,129,161)
' 245	Cyan vivid		#00687d	RGB(0,104,125)
' 246	Cyan vivid		#0e4f5c	RGB(14,79,92)
' 247	Cyan vivid		#093b44	RGB(9,59,68)
' 248	Blue cool		#e7f2f5	RGB(231,242,245)
' 249	Blue cool		#dae9ee	RGB(218,233,238)
' 250	Blue cool		#adcfdc	RGB(173,207,220)
' 251	Blue cool		#82b4c9	RGB(130,180,201)
' 252	Blue cool		#6499af	RGB(100,153,175)
' 253	Blue cool		#3a7d95	RGB(58,125,149)
' 254	Blue cool		#2e6276	RGB(46,98,118)
' 255	Blue cool		#224a58	RGB(34,74,88)
' 256	Blue cool		#14333d	RGB(20,51,61)
' 257	Blue cool		#0f191c	RGB(15,25,28)
' 258	Blue cool vivid		#e1f3f8	RGB(225,243,248)
' 259	Blue cool vivid		#c3ebfa	RGB(195,235,250)
' 260	Blue cool vivid		#97d4ea	RGB(151,212,234)
' 261	Blue cool vivid		#59b9de	RGB(89,185,222)
' 262	Blue cool vivid		#28a0cb	RGB(40,160,203)
' 263	Blue cool vivid		#0d7ea2	RGB(13,126,162)
' 264	Blue cool vivid		#07648d	RGB(7,100,141)
' 265	Blue cool vivid		#074b69	RGB(7,75,105)
' 266	Blue cool vivid		#002d3f	RGB(0,45,63)
' 267	Blue		#eff6fb	RGB(239,246,251)
' 268	Blue		#d9e8f6	RGB(217,232,246)
' 269	Blue		#aacdec	RGB(170,205,236)
' 270	Blue		#73b3e7	RGB(115,179,231)
' 271	Blue		#4f97d1	RGB(79,151,209)
' 272	Blue		#2378c3	RGB(35,120,195)
' 273	Blue		#2c608a	RGB(44,96,138)
' 274	Blue		#274863	RGB(39,72,99)
' 275	Blue		#1f303e	RGB(31,48,62)
' 276	Blue		#11181d	RGB(17,24,29)
' 277	Blue vivid		#e8f5ff	RGB(232,245,255)
' 278	Blue vivid		#cfe8ff	RGB(207,232,255)
' 279	Blue vivid		#a1d3ff	RGB(161,211,255)
' 280	Blue vivid		#58b4ff	RGB(88,180,255)
' 281	Blue vivid		#2491ff	RGB(36,145,255)
' 282	Blue vivid		#0076d6	RGB(0,118,214)
' 283	Blue vivid		#005ea2	RGB(0,94,162)
' 284	Blue vivid		#0b4778	RGB(11,71,120)
' 285	Blue vivid		#112f4e	RGB(17,47,78)
' 286	Blue warm		#ecf1f7	RGB(236,241,247)
' 287	Blue warm		#e1e7f1	RGB(225,231,241)
' 288	Blue warm		#bbcae4	RGB(187,202,228)
' 289	Blue warm		#98afd2	RGB(152,175,210)
' 290	Blue warm		#7292c7	RGB(114,146,199)
' 291	Blue warm		#4a77b4	RGB(74,119,180)
' 292	Blue warm		#345d96	RGB(52,93,150)
' 293	Blue warm		#2f4668	RGB(47,70,104)
' 294	Blue warm		#252f3e	RGB(37,47,62)
' 295	Blue warm		#13171f	RGB(19,23,31)
' 296	Blue warm vivid		#edf5ff	RGB(237,245,255)
' 297	Blue warm vivid		#d4e5ff	RGB(212,229,255)
' 298	Blue warm vivid		#adcdff	RGB(173,205,255)
' 299	Blue warm vivid		#81aefc	RGB(129,174,252)
' 300	Blue warm vivid		#5994f6	RGB(89,148,246)
' 301	Blue warm vivid		#2672de	RGB(38,114,222)
' 302	Blue warm vivid		#0050d8	RGB(0,80,216)
' 303	Blue warm vivid		#1a4480	RGB(26,68,128)
' 304	Blue warm vivid		#162e51	RGB(22,46,81)
' 305	Indigo cool		#eef0f9	RGB(238,240,249)
' 306	Indigo cool		#e1e6f9	RGB(225,230,249)
' 307	Indigo cool		#bbc8f5	RGB(187,200,245)
' 308	Indigo cool		#96abee	RGB(150,171,238)
' 309	Indigo cool		#6b8ee8	RGB(107,142,232)
' 310	Indigo cool		#496fd8	RGB(73,111,216)
' 311	Indigo cool		#3f57a6	RGB(63,87,166)
' 312	Indigo cool		#374274	RGB(55,66,116)
' 313	Indigo cool		#292d42	RGB(41,45,66)
' 314	Indigo cool		#151622	RGB(21,22,34)
' 315	Indigo cool vivid		#edf0ff	RGB(237,240,255)
' 316	Indigo cool vivid		#dee5ff	RGB(222,229,255)
' 317	Indigo cool vivid		#b8c8ff	RGB(184,200,255)
' 318	Indigo cool vivid		#94adff	RGB(148,173,255)
' 319	Indigo cool vivid		#628ef4	RGB(98,142,244)
' 320	Indigo cool vivid		#4866ff	RGB(72,102,255)
' 321	Indigo cool vivid		#3e4ded	RGB(62,77,237)
' 322	Indigo cool vivid		#222fbf	RGB(34,47,191)
' 323	Indigo cool vivid		#1b2b85	RGB(27,43,133)
' 324	Indigo		#efeff8	RGB(239,239,248)
' 325	Indigo		#e5e4fa	RGB(229,228,250)
' 326	Indigo		#c5c5f3	RGB(197,197,243)
' 327	Indigo		#a5a8eb	RGB(165,168,235)
' 328	Indigo		#8889db	RGB(136,137,219)
' 329	Indigo		#676cc8	RGB(103,108,200)
' 330	Indigo		#4d52af	RGB(77,82,175)
' 331	Indigo		#3d4076	RGB(61,64,118)
' 332	Indigo		#2b2c40	RGB(43,44,64)
' 333	Indigo		#16171f	RGB(22,23,31)
' 334	Indigo vivid		#f0f0ff	RGB(240,240,255)
' 335	Indigo vivid		#e0e0ff	RGB(224,224,255)
' 336	Indigo vivid		#ccceff	RGB(204,206,255)
' 337	Indigo vivid		#a3a7fa	RGB(163,167,250)
' 338	Indigo vivid		#8289ff	RGB(130,137,255)
' 339	Indigo vivid		#656bd7	RGB(101,107,215)
' 340	Indigo vivid		#4a50c4	RGB(74,80,196)
' 341	Indigo vivid		#3333a3	RGB(51,51,163)
' 342	Indigo vivid		#212463	RGB(33,36,99)
' 343	Indigo warm		#f1eff7	RGB(241,239,247)
' 344	Indigo warm		#e7e3fa	RGB(231,227,250)
' 345	Indigo warm		#cbc4f2	RGB(203,196,242)
' 346	Indigo warm		#afa5e8	RGB(175,165,232)
' 347	Indigo warm		#9287d8	RGB(146,135,216)
' 348	Indigo warm		#7665d1	RGB(118,101,209)
' 349	Indigo warm		#5e519e	RGB(94,81,158)
' 350	Indigo warm		#453c7b	RGB(69,60,123)
' 351	Indigo warm		#2e2c40	RGB(46,44,64)
' 352	Indigo warm		#18161d	RGB(24,22,29)
' 353	Indigo warm vivid		#f5f2ff	RGB(245,242,255)
' 354	Indigo warm vivid		#e4deff	RGB(228,222,255)
' 355	Indigo warm vivid		#cfc4fd	RGB(207,196,253)
' 356	Indigo warm vivid		#b69fff	RGB(182,159,255)
' 357	Indigo warm vivid		#967efb	RGB(150,126,251)
' 358	Indigo warm vivid		#745fe9	RGB(116,95,233)
' 359	Indigo warm vivid		#5942d2	RGB(89,66,210)
' 360	Indigo warm vivid		#3d2c9d	RGB(61,44,157)
' 361	Indigo warm vivid		#261f5b	RGB(38,31,91)
' 362	Violet		#f4f1f9	RGB(244,241,249)
' 363	Violet		#ebe3f9	RGB(235,227,249)
' 364	Violet		#d0c3e9	RGB(208,195,233)
' 365	Violet		#b8a2e3	RGB(184,162,227)
' 366	Violet		#9d84d2	RGB(157,132,210)
' 367	Violet		#8168b3	RGB(129,104,179)
' 368	Violet		#665190	RGB(102,81,144)
' 369	Violet		#4c3d69	RGB(76,61,105)
' 370	Violet		#312b3f	RGB(49,43,63)
' 371	Violet		#18161d	RGB(24,22,29)
' 372	Violet vivid		#f7f2ff	RGB(247,242,255)
' 373	Violet vivid		#ede3ff	RGB(237,227,255)
' 374	Violet vivid		#d5bfff	RGB(213,191,255)
' 375	Violet vivid		#c39deb	RGB(195,157,235)
' 376	Violet vivid		#ad79e9	RGB(173,121,233)
' 377	Violet vivid		#9355dc	RGB(147,85,220)
' 378	Violet vivid		#783cb9	RGB(120,60,185)
' 379	Violet vivid		#54278f	RGB(84,39,143)
' 380	Violet vivid		#39215e	RGB(57,33,94)
' 381	Violet warm		#f8f0f9	RGB(248,240,249)
' 382	Violet warm		#f6dff8	RGB(246,223,248)
' 383	Violet warm		#e2bee4	RGB(226,190,228)
' 384	Violet warm		#d29ad8	RGB(210,154,216)
' 385	Violet warm		#bf77c8	RGB(191,119,200)
' 386	Violet warm		#b04abd	RGB(176,74,189)
' 387	Violet warm		#864381	RGB(134,67,129)
' 388	Violet warm		#5c395a	RGB(92,57,90)
' 389	Violet warm		#382936	RGB(56,41,54)
' 390	Violet warm		#1b151b	RGB(27,21,27)
' 391	Violet warm vivid		#fef2ff	RGB(254,242,255)
' 392	Violet warm vivid		#fbdcff	RGB(251,220,255)
' 393	Violet warm vivid		#f4b2ff	RGB(244,178,255)
' 394	Violet warm vivid		#ee83ff	RGB(238,131,255)
' 395	Violet warm vivid		#d85bef	RGB(216,91,239)
' 396	Violet warm vivid		#be32d0	RGB(190,50,208)
' 397	Violet warm vivid		#93348c	RGB(147,52,140)
' 398	Violet warm vivid		#711e6c	RGB(113,30,108)
' 399	Violet warm vivid		#481441	RGB(72,20,65)
' 400	Magenta		#f9f0f2	RGB(249,240,242)
' 401	Magenta		#f6e1e8	RGB(246,225,232)
' 402	Magenta		#f0bbcc	RGB(240,187,204)
' 403	Magenta		#e895b3	RGB(232,149,179)
' 404	Magenta		#e0699f	RGB(224,105,159)
' 405	Magenta		#c84281	RGB(200,66,129)
' 406	Magenta		#8b4566	RGB(139,69,102)
' 407	Magenta		#66364b	RGB(102,54,75)
' 408	Magenta		#402731	RGB(64,39,49)
' 409	Magenta		#1b1617	RGB(27,22,23)
' 410	Magenta vivid		#fff2f5	RGB(255,242,245)
' 411	Magenta vivid		#ffddea	RGB(255,221,234)
' 412	Magenta vivid		#ffb4cf	RGB(255,180,207)
' 413	Magenta vivid		#ff87b2	RGB(255,135,178)
' 414	Magenta vivid		#fd4496	RGB(253,68,150)
' 415	Magenta vivid		#d72d79	RGB(215,45,121)
' 416	Magenta vivid		#ab2165	RGB(171,33,101)
' 417	Magenta vivid		#731f44	RGB(115,31,68)
' 418	Magenta vivid		#4f172e	RGB(79,23,46)
' 419	Gray cool		#fbfcfd	RGB(251,252,253)
' 420	Gray cool		#f7f9fa	RGB(247,249,250)
' 421	Gray cool		#f5f6f7	RGB(245,246,247)
' 422	Gray cool		#f1f3f6	RGB(241,243,246)
' 423	Gray cool		#edeff0	RGB(237,239,240)
' 424	Gray cool		#dfe1e2	RGB(223,225,226)
' 425	Gray cool		#c6cace	RGB(198,202,206)
' 426	Gray cool		#a9aeb1	RGB(169,174,177)
' 427	Gray cool		#8d9297	RGB(141,146,151)
' 428	Gray cool		#71767a	RGB(113,118,122)
' 429	Gray cool		#565c65	RGB(86,92,101)
' 430	Gray cool		#3d4551	RGB(61,69,81)
' 431	Gray cool		#2d2e2f	RGB(45,46,47)
' 432	Gray cool		#1c1d1f	RGB(28,29,31)
' 433	Gray		#fcfcfc	RGB(252,252,252)
' 434	Gray		#f9f9f9	RGB(249,249,249)
' 435	Gray		#f6f6f6	RGB(246,246,246)
' 436	Gray	    #f3f3f3	    RGB(243,243,243)
' 437	Gray		#f0f0f0	RGB(240,240,240)
' 438	Gray		#e6e6e6	RGB(230,230,230)
' 439	Gray		#c9c9c9	RGB(201,201,201)
' 440	Gray		#adadad	RGB(173,173,173)
' 441	Gray		#919191	RGB(145,145,145)
' 442	Gray		#757575	RGB(117,117,117)
' 443	Gray		#5c5c5c	RGB(92,92,92)
' 444	Gray		#454545	RGB(69,69,69)
' 445	Gray		#2e2e2e	RGB(46,46,46)
' 446	Gray		#1b1b1b	RGB(27,27,27)
' 447	Gray warm		#fcfcfb	RGB(252,252,251)
' 448	Gray warm		#f9f9f7	RGB(249,249,247)
' 449	Gray warm		#f6f6f2	RGB(246,246,242)
' 450	Gray warm		#f5f5f0	RGB(245,245,240)
' 451	Gray warm		#f0f0ec	RGB(240,240,236)
' 452	Gray warm		#e6e6e2	RGB(230,230,226)
' 453	Gray warm		#cac9c0	RGB(202,201,192)
' 454	Gray warm		#afaea2	RGB(175,174,162)
' 455	Gray warm		#929285	RGB(146,146,133)
' 456	Gray warm		#76766a	RGB(118,118,106)
' 457	Gray warm		#5d5d52	RGB(93,93,82)
' 458	Gray warm		#454540	RGB(69,69,64)
' 459	Gray warm		#2e2e2a	RGB(46,46,42)
' 460	Gray warm		#171716	RGB(23,23,22)
