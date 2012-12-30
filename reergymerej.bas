Attribute VB_Name = "reergymerej"
Private Function dictionary()
    Dim d(1000) As String
    
    '============================================================================================
    '   ONLY EDIT BELOW THIS LINE
    '
    '   "THIS_PART:xx:xxxx" is what you want the value to be.
    '   "xxx:THESE PARTS:AND THESE" are what you want to be replaced.
    
    d(0) = "AFG:AF:Afghanistan"
    d(1) = "ALB:AL:Albania"
    d(2) = "DZA:DZ:Algeria"
    d(3) = "ASM:AS:American Samoa"
    d(4) = "AND:AD:Andorra"
    d(5) = "AGO:AO:Angola"
    d(6) = "AIA:AI:Anguilla"
    d(7) = "ATA:AQ:Antarctica"
    d(8) = "ATG:AG:Antigua and Barbuda"
    d(9) = "ARG:AR:Argentina"
    d(10) = "ARM:AM:Armenia"
    d(11) = "ABW:AW:Aruba"
    d(12) = "AUS:AU:Australia"
    d(13) = "AUT:AT:Austria"
    d(14) = "AZE:AZ:Azerbaijan"
    d(15) = "BHS:BS:Bahamas"
    d(16) = "BHR:BH:Bahrain"
    d(17) = "BGD:BD:Bangladesh"
    d(18) = "BRB:BB:Barbados"
    d(19) = "BLR:BY:Belarus"
    d(20) = "BEL:BE:Belgium"
    d(21) = "BLZ:BZ:Belize"
    d(22) = "BEN:BJ:Benin"
    d(23) = "BMU:BM:Bermuda"
    d(24) = "BTN:BT:Bhutan"
    d(25) = "BOL:BO:Bolivia"
    d(26) = "BIH:BA:Bosnia and Herzegovina"
    d(27) = "BWA:BW:Botswana"
    d(28) = "BRA:BR:Brazil"
    d(29) = "IOT:IO:British Indian Ocean Territory"
    d(30) = "VGB:VG:British Virgin Islands"
    d(31) = "BRN:BN:Brunei"
    d(32) = "BGR:BG:Bulgaria"
    d(33) = "BFA:BF:Burkina Faso"
    d(34) = "MMR:MM:Burma (Myanmar)"
    d(35) = "BDI:BI:Burundi"
    d(36) = "KHM:KH:Cambodia"
    d(37) = "CMR:CM:Cameroon"
    d(38) = "CAN:CA:Canada"
    d(39) = "CPV:CV:Cape Verde"
    d(40) = "CYM:KY:Cayman Islands"
    d(41) = "CAF:CF:Central African Republic"
    d(42) = "TCD:TD:Chad"
    d(43) = "CHL:CL:Chile"
    d(44) = "CHN:CN:China"
    d(45) = "CXR:CX:Christmas Island"
    d(46) = "CCK:CC:Cocos (Keeling) Islands"
    d(47) = "COL:CO:Colombia"
    d(48) = "COM:KM:Comoros"
    d(49) = "COK:CK:Cook Islands"
    d(50) = "CRC:CR:Costa Rica"
    d(51) = "HRV:HR:Croatia"
    d(52) = "CUB:CU:Cuba"
    d(53) = "CYP:CY:Cyprus"
    d(54) = "CZE:CZ:Czech Republic"
    d(55) = "COD:CD:Democratic Republic of the Congo"
    d(56) = "DNK:DK:Denmark"
    d(57) = "DJI:DJ:Djibouti"
    d(58) = "DMA:DM:Dominica"
    d(59) = "DOM:DO:Dominican Republic"
    d(60) = "ECU:EC:Ecuador"
    d(61) = "EGY:EG:Egypt"
    d(62) = "SLV:SV:El Salvador"
    d(63) = "GNQ:GQ:Equatorial Guinea"
    d(64) = "ERI:ER:Eritrea"
    d(65) = "EST:EE:Estonia"
    d(66) = "ETH:ET:Ethiopia"
    d(67) = "FLK:FK:Falkland Islands"
    d(68) = "FRO:FO:Faroe Islands"
    d(69) = "FJI:FJ:Fiji"
    d(70) = "FIN:FI:Finland"
    d(71) = "FRA:FR:France"
    d(72) = "PYF:PF:French Polynesia"
    d(73) = "GAB:GA:Gabon"
    d(74) = "GMB:GM:Gambia"
    d(75) = "GEO:GE:Georgia"
    d(76) = "DEU:DE:Germany"
    d(77) = "GHA:GH:Ghana"
    d(78) = "GIB:GI:Gibraltar"
    d(79) = "GRC:GR:Greece"
    d(80) = "GRL:GL:Greenland"
    d(81) = "GRD:GD:Grenada"
    d(82) = "GUM:GU:Guam"
    d(83) = "GTM:GT:Guatemala"
    d(84) = "GIN:GN:Guinea"
    d(85) = "GNB:GW:Guinea-Bissau"
    d(86) = "GUY:GY:Guyana"
    d(87) = "HTI:HT:Haiti"
    d(88) = "VAT:VA:Holy See (Vatican City)"
    d(89) = "HND:HN:Honduras"
    d(90) = "HKG:HK:Hong Kong"
    d(91) = "HUN:HU:Hungary"
    d(92) = "IS:IS:Iceland"
    d(93) = "IND:IN:India"
    d(94) = "IDN:ID:Indonesia"
    d(95) = "IRN:IR:Iran"
    d(96) = "IRQ:IQ:Iraq"
    d(97) = "IRL:IE:Ireland"
    d(98) = "IMN:IM:Isle of Man"
    d(99) = "ISR:IL:Israel"
    d(100) = "ITA:IT:Italy"
    d(101) = "CIV:CI:Ivory Coast"
    d(102) = "JAM:JM:Jamaica"
    d(103) = "JPN:JP:Japan"
    d(104) = "JEY:JE:Jersey"
    d(105) = "JOR:JO:Jordan"
    d(106) = "KAZ:KZ:Kazakhstan"
    d(107) = "KEN:KE:Kenya"
    d(108) = "KIR:KI:Kiribati"
    d(109) = "KWT:KW:Kuwait"
    d(110) = "KGZ:KG:Kyrgyzstan"
    d(111) = "LAO:LA:Laos"
    d(112) = "LVA:LV:Latvia"
    d(113) = "LBN:LB:Lebanon"
    d(114) = "LSO:LS:Lesotho"
    d(115) = "LBR:LR:Liberia"
    d(116) = "LBY:LY:Libya"
    d(117) = "LIE:LI:Liechtenstein"
    d(118) = "LTU:LT:Lithuania"
    d(119) = "LUX:LU:Luxembourg"
    d(120) = "MAC:MO:Macau"
    d(121) = "MKD:MK:Macedonia"
    d(122) = "MDG:MG:Madagascar"
    d(123) = "MWI:MW:Malawi"
    d(124) = "MYS:MY:Malaysia"
    d(125) = "MDV:MV:Maldives"
    d(126) = "MLI:ML:Mali"
    d(127) = "MLT:MT:Malta"
    d(128) = "MHL:MH:Marshall Islands"
    d(129) = "MRT:MR:Mauritania"
    d(130) = "MUS:MU:Mauritius"
    d(131) = "MYT:YT:Mayotte"
    d(132) = "MEX:MX:Mexico"
    d(133) = "FSM:FM:Micronesia"
    d(134) = "MDA:MD:Moldova"
    d(135) = "MCO:MC:Monaco"
    d(136) = "MNG:MN:Mongolia"
    d(137) = "MNE:ME:Montenegro"
    d(138) = "MSR:MS:Montserrat"
    d(139) = "MAR:MA:Morocco"
    d(140) = "MOZ:MZ:Mozambique"
    d(141) = "NAM:NA:Namibia"
    d(142) = "NRU:NR:Nauru"
    d(143) = "NPL:NP:Nepal"
    d(144) = "NLD:NL:Netherlands"
    d(145) = "ANT:AN:Netherlands Antilles"
    d(146) = "NCL:NC:New Caledonia"
    d(147) = "NZL:NZ:New Zealand"
    d(148) = "NIC:NI:Nicaragua"
    d(149) = "NER:NE:Niger"
    d(150) = "NGA:NG:Nigeria"
    d(151) = "NIU:NU:Niue"
    d(152) = "NFK:Norfolk Island"
    d(153) = "PRK:KP:North Korea"
    d(154) = "MNP:MP:Northern Mariana Islands"
    d(155) = "NOR:NO:Norway"
    d(156) = "OMN:OM:Oman"
    d(157) = "PAK:PK:Pakistan"
    d(158) = "PLW:PW:Palau"
    d(159) = "PAN:PA:Panama"
    d(160) = "PNG:PG:Papua New Guinea"
    d(161) = "PRY:PY:Paraguay"
    d(162) = "PER:PE:Peru"
    d(163) = "PHL:PH:Philippines"
    d(164) = "PCN:PN:Pitcairn Islands"
    d(165) = "POL:PL:Poland"
    d(166) = "PRT:PT:Portugal"
    d(167) = "PRI:PR:Puerto Rico"
    d(168) = "QAT:QA:Qatar"
    d(169) = "COG:CG:Republic of the Congo"
    d(170) = "ROU:RO:Romania"
    d(171) = "RUS:RU:Russia"
    d(172) = "RWA:RW:Rwanda"
    d(173) = "BLM:BL:Saint Barthelemy"
    d(174) = "SHN:SH:Saint Helena"
    d(175) = "KNA:KN:Saint Kitts and Nevis"
    d(176) = "LCA:LC:Saint Lucia"
    d(177) = "MAF:MF:Saint Martin"
    d(178) = "SPM:PM:Saint Pierre and Miquelon"
    d(179) = "VCT:VC:Saint Vincent and the Grenadines"
    d(180) = "WSM:WS:Samoa"
    d(181) = "SMR:SM:San Marino"
    d(182) = "STP:ST:Sao Tome and Principe"
    d(183) = "SAU:SA:Saudi Arabia"
    d(184) = "SEN:SN:Senegal"
    d(185) = "SRB:RS:Serbia"
    d(186) = "SYC:SC:Seychelles"
    d(187) = "SLE:SL:Sierra Leone"
    d(188) = "SGP:SG:Singapore"
    d(189) = "SVK:SK:Slovakia"
    d(190) = "SVN:SI:Slovenia"
    d(191) = "SLB:SB:Solomon Islands"
    d(192) = "SOM:SO:Somalia"
    d(193) = "ZAF:ZA:South Africa"
    d(194) = "KOR:KR:South Korea"
    d(195) = "ESP:ES:Spain"
    d(196) = "LKA:LK:Sri Lanka"
    d(197) = "SDN:SD:Sudan"
    d(198) = "SUR:SR:Suriname"
    d(199) = "SJM:SJ:Svalbard"
    d(200) = "SWZ:SZ:Swaziland"
    d(201) = "SWE:SE:Sweden"
    d(202) = "CHE:CH:Switzerland"
    d(203) = "SYR:SY:Syria"
    d(204) = "TWN:TW:Taiwan"
    d(205) = "TJK:TJ:Tajikistan"
    d(206) = "TZA:TZ:Tanzania"
    d(207) = "THA:TH:Thailand"
    d(208) = "TLS:TL:Timor-Leste"
    d(209) = "TGO:TG:Togo"
    d(210) = "TKL:TK:Tokelau"
    d(211) = "TON:TO:Tonga"
    d(212) = "TTO:TT:Trinidad and Tobago"
    d(213) = "TUN:TN:Tunisia"
    d(214) = "TUR:TR:Turkey"
    d(215) = "TKM:TM:Turkmenistan"
    d(216) = "TCA:TC:Turks and Caicos Islands"
    d(217) = "TUV:TV:Tuvalu"
    d(218) = "UGA:UG:Uganda"
    d(219) = "UKR:UA:Ukraine"
    d(220) = "ARE:AE:United Arab Emirates"
    d(221) = "GBR:GB:United Kingdom"
    d(222) = "USA:US:United States:America"
    d(223) = "URY:UY:Uruguay"
    d(224) = "VIR:VI:US Virgin Islands"
    d(225) = "UZB:UZ:Uzbekistan"
    d(226) = "VUT:VU:Vanuatu"
    d(227) = "VEN:VE:Venezuela"
    d(228) = "VNM:VN:Vietnam"
    d(229) = "WLF:WF:Wallis and Futuna"
    d(230) = "ESH:EH:Western Sahara"
    d(231) = "YEM:YE:Yemen"
    d(232) = "ZMB:ZM:Zambia"
    d(233) = "ZWE:ZW:Zimbabwe"
    
    
    '   ONLY EDIT ABOVE THIS LINE
    '============================================================================================
    
    dictionary = d
End Function

Sub I_love_Amanda()
Attribute I_love_Amanda.VB_Description = "Do a whole bunch of find/replace operations at once."
Attribute I_love_Amanda.VB_ProcData.VB_Invoke_Func = "i\n14"

    Application.ScreenUpdating = False
    
    Dim delimiter As String
    delimiter = ":"

    '   remember active cell
    Dim a As Range
    Set a = activeCell
    
    Dim vals
    vals = dictionary()
    
    Dim v() As String
    
    '   select the column
    a.EntireColumn.Select
    
    For i = LBound(vals) To UBound(vals)
        If vals(i) <> "" Then
                
            v = Split(vals(i), delimiter)
            
            For j = 1 To UBound(v)
            
                '   replace
                Selection.Replace What:=v(j), Replacement:=v(0), LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False
            
            Next j
            
            '   final replacement for case insensitivity
            Selection.Replace What:=v(0), Replacement:=v(0), LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
        End If
    Next i
    
    '   restore original selection
    a.Select
    
    Application.ScreenUpdating = True
End Sub





