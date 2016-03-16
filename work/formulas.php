<?php

/**
 * Ну вы понели :)
 * Это формулы из жёлтых ячеек. Теперь нужно сделать три вещи:
 * - Заменить обозначения диапозонов на сумму простого перечисления его ячеек
 * - Перегнать буквенные обозначения стобцов A, B, AC... в числовые индексы стобцов в нашей объединённой матрице
 * - Заменить индексы рядов из xls на индексы в нашей матрице. Или же в ней так исправить индексирование рядов (вряд ли)
 */
$formulas =
	[
		'8'    => ['B8' => 'SUM($I2150:$J2169)+SUM($I2175:$J2195)', 'C8' => 'SUM($U2150:$V2169,$U2175:$V2195)+SUM($AC2150:$AD2169,$AC2175:$AD2195)', 'E8' => 'SUM($I2150:$I2169)+SUM($I2173:$I2199)', 'F8' => 'SUM($U2150:$U2169,$U2173:$U2199)+SUM($AC2150:$AC2169,$AC2175:$AC2196)', ],
		'9'    => ['B9' => 'SUM($I2171:$J2171)', 'C9' => 'SUM($U2171:$V2171,$AC2171:$AD2171)', 'E9' => 'SUM($I2171:$I2172)', 'F9' => 'SUM($U2171:$U2172,$AC2171:$AC2172)', ],
		'10'   => ['B10' => 'SUM($I2295:$J2295,$I2245:$J2250)', 'C10' => 'SUM($U2295:$V2295,$U2245:$V2250,$AC2245:$AD2250)', 'E10' => 'SUM($I2295,$I2245:$I2250)', 'F10' => 'SUM($U2295,$U2245:$U2250,$AC2245:$AC2250)', ],
		'11'   => ['B11' => 'SUM($I184:$J184)', 'C11' => '$U184+$V184+$AC184+$AD184', 'E11' => 'SUM($I184:$I184)', 'F11' => '$U184+$AC184', ],
		'12'   => ['B12' => 'SUM($I54:$J56)', 'C12' => 'SUM($U54:$V56,$AC54:$AD56)', 'E12' => 'SUM($I54:$I56)', 'F12' => 'SUM($U54:$U56,$AC54:$AC56)', ],
		'13'   => ['B13' => 'SUM($J187:$J187,$J1494:$J1494)', 'C13' => '$V187+$AD187+$V1494+$AD1494', 'E13' => 'SUM($I187:$I187,$I1494:$I1494,$I1295:$I1295)', 'F13' => '$U187+$AC187+$U1494+$AC1494+$U1295+$AC1295', ],
		'14'   => ['B14' => 'SUM($J1308)', 'C14' => 'SUM($V1308,$AD1308)', 'E14' => 'SUM($I1308)+I1', 'F14' => 'SUM($U1308,$AC1308)', ],
		'15'   => ['B15' => 'SUM($I2294:$J2294)', 'C15' => '$U2294+$V2294', 'E15' => 'SUM($I2294,$I1767,$I1294)', 'F15' => '$U2294+$U1767+$AC1767+$U1294+$AC1294', ],
		'16'   => ['B16' => 'SUM($I57:$J98)', 'C16' => 'SUM($U57:$V98,$AC57:$AD98)', 'E16' => 'SUM($I57:$I98)+$I2284', 'F16' => 'SUM($U57:$U98,$AC57:$AC98)+$U2284+$AC2284', ],
		'17'   => ['B17' => 'SUM($I52:$J53)', 'C17' => '$U52+$V52+$AC52+$AD52+$U53+$V53+$AC53+$AD53', 'E17' => 'SUM($I52:$I53)', 'F17' => '$U52+$AC52+$U53+$AC53', ],
		'18'   => ['B18' => 'SUM($I2252:$J2289)', 'C18' => 'SUM($AC2252:$AD2289,$U2252:$V2289)', 'E18' => 'SUM($I2252:$I2279,$I2281:$I2284,$I2285:$I2289)', 'F18' => 'SUM($AC2252:$AC2282,$AC2283:$AC2284,$AC2285:$AC2289,$U2252:$U2282,$U2283:$U2284,$U2285:$U2289)', ],
		'19'   => ['B19' => 'SUM($I9:$J40)', 'C19' => 'SUM($AC9:$AD40,$U9:$V40)', 'E19' => 'SUM($I8:$I47,$I1292:$I1292)', 'F19' => 'SUM($AC8:$AC47,$U8:$U47)', ],
		'20'   => ['B20' => '', 'C20' => '', 'E20' => '$I126', 'F20' => '$U126+$AC126', ],
		'21'   => ['B21' => '', 'C21' => '', 'E21' => '$I127', 'F21' => '$U127+$AC127', ],
		'22'   => ['B22' => '', 'C22' => '', 'E22' => 'I117', 'F22' => 'U117+AC117', ],
		'23'   => ['B23' => '', 'C23' => '', 'E23' => 'I118', 'F23' => 'U118+AC118', ],
		'24'   => ['B24' => 'SUM($J1604:$J1604)', 'C24' => 'SUM($V1604:$V1604,$AD1604:$AD1604)', 'E24' => 'SUM($I1604)', 'F24' => 'SUM($U1604,$AC1604)', ],
		'25'   => ['B25' => '', 'C25' => '', 'E25' => '$I976', 'F25' => '$U976+$AC976', ],
		'26'   => ['B26' => '', 'C26' => '', 'E26' => '$I1857', 'F26' => '$U1857+$AC1857', ],
		'27'   => ['B27' => '', 'C27' => '', 'E27' => '$I933', 'F27' => '$U933+$AC933', ],
		'28'   => ['B28' => '', 'C28' => '', 'E28' => '$I935', 'F28' => '$U935+$AC935', ],
		'29'   => ['B29' => '', 'C29' => '', 'E29' => '$I1585', 'F29' => '$U1585+$AC1585', ],
		'30'   => ['B30' => '', 'C30' => '', 'E30' => '$I336', 'F30' => '$U336+$AC336', ],
		'31'   => ['B31' => '', 'C31' => '', 'E31' => '$I1586', 'F31' => '$U1586+$AC1586', ],
		'32'   => ['B32' => '', 'C32' => '', 'E32' => '$I1698', 'F32' => '$U1698+$AC1698', ],
		'33'   => ['B33' => '', 'C33' => '', 'E33' => '$I1706', 'F33' => '$U1706+$AC1706', ],
		'34'   => ['B34' => '', 'C34' => '', 'E34' => '$I1825', 'F34' => '$U1825+$AC1825', ],
		'35'   => ['B35' => '', 'C35' => '', 'E35' => '$I1584', 'F35' => '$U1584+$AC1584', ],
		'36'   => ['B36' => '', 'C36' => '', 'E36' => '$I941', 'F36' => '$U941+$AC941', ],
		'37'   => ['B37' => '', 'C37' => '', 'E37' => '$I1888', 'F37' => '$U1888+$AC1888', ],
		'38'   => ['B38' => '', 'C38' => '', 'E38' => '$I1942', 'F38' => '$U1942+$AC1942', ],
		'39'   => ['B39' => '', 'C39' => '', 'E39' => '$I230', 'F39' => '$U230+$AC230', ],
		'40'   => ['B40' => '', 'C40' => '', 'E40' => '$I239', 'F40' => '$U239+$AC239', ],
		'41'   => ['B41' => '', 'C41' => '', 'E41' => '$I650', 'F41' => '$U650+$AC650', ],
		'42'   => ['B42' => '', 'C42' => '', 'E42' => '$I1320', 'F42' => '$U1320+$AC1320', ],
		'43'   => ['B43' => '', 'C43' => '', 'E43' => 'I1471', 'F43' => 'U1471+AC1471', ],
		'44'   => ['B44' => '', 'C44' => '', 'E44' => 'I1671', 'F44' => 'U1671+AC1671', ],
		'45'   => ['B45' => '', 'C45' => '', 'E45' => 'I341', 'F45' => 'U341+AC341', ],
		'46'   => ['B46' => '', 'C46' => '', 'E46' => 'I1226', 'F46' => 'U1226+AC1226', ],
		'47'   => ['B47' => '', 'C47' => '', 'E47' => 'I1636', 'F47' => 'U1636+AC1636', ],
		'48'   => ['B48' => '', 'C48' => '', 'E48' => 'I1991', 'F48' => 'U1991+AC1991', ],
		'49'   => ['B49' => '', 'C49' => '', 'E49' => '', 'F49' => '', ],
		'50'   => ['B50' => '', 'C50' => '', 'E50' => '', 'F50' => '', ],
		'51'   => ['B51' => '', 'C51' => '', 'E51' => '', 'F51' => '', ],
		'52'   => ['B52' => '', 'C52' => '', 'E52' => '', 'F52' => '', ],
		'53'   => ['B53' => '', 'C53' => '', 'E53' => '', 'F53' => '', ],
		'54'   => ['B54' => '', 'C54' => '', 'E54' => '', 'F54' => '', ],
		'55'   => ['B55' => '', 'C55' => '', 'E55' => '', 'F55' => '', ],
		'56'   => ['B56' => '', 'C56' => '', 'E56' => '', 'F56' => '', ],
		'57'   => ['B57' => '', 'C57' => '', 'E57' => '', 'F57' => '', ],
		'58'   => ['B58' => '', 'C58' => '', 'E58' => 'I2014', 'F58' => 'U2014+AC2014', ],
		'59'   => ['B59' => '', 'C59' => '', 'E59' => '', 'F59' => '', ],
		'60'   => ['B60' => '', 'C60' => '', 'E60' => '', 'F60' => '', ],
		'61'   => ['B61' => '', 'C61' => '', 'E61' => '', 'F61' => '', ],
		'62'   => ['B62' => '', 'C62' => '', 'E62' => '', 'F62' => '', ],
		'63'   => ['B63' => '', 'C63' => '', 'E63' => '', 'F63' => '', ],
		'64'   => ['B64' => '', 'C64' => '', 'E64' => '', 'F64' => '', ],
		'65'   => ['B65' => '', 'C65' => '', 'E65' => '', 'F65' => '', ],
		'66'   => ['B66' => '', 'C66' => '', 'E66' => 'I2027', 'F66' => 'U2027+AC2027', ],
		'67'   => ['B67' => '', 'C67' => '', 'E67' => 'I2028', 'F67' => 'U2028+AC2028', ],
		'68'   => ['B68' => '', 'C68' => '', 'E68' => 'I2029', 'F68' => 'U2029+AC2029', ],
		'69'   => ['B69' => '', 'C69' => '', 'E69' => 'I2030', 'F69' => 'U2030+AC2030', ],
		'70'   => ['B70' => '', 'C70' => '', 'E70' => 'I2051', 'F70' => 'U2051+AC2051', ],
		'71'   => ['B71' => '', 'C71' => '', 'E71' => '', 'F71' => '', ],
		'72'   => ['B72' => '', 'C72' => '', 'E72' => '', 'F72' => '', ],
		'73'   => ['B73' => '', 'C73' => '', 'E73' => '', 'F73' => '', ],
		'74'   => ['B74' => '', 'C74' => '', 'E74' => '', 'F74' => '', ],
		'75'   => ['B75' => '', 'C75' => '', 'E75' => '', 'F75' => '', ],
		'76'   => ['B76' => '', 'C76' => '', 'E76' => '', 'F76' => '', ],
		'77'   => ['B77' => '', 'C77' => '', 'E77' => '', 'F77' => '', ],
		'78'   => ['B78' => '', 'C78' => '', 'E78' => '', 'F78' => '', ],
		'79'   => ['B79' => '', 'C79' => '', 'E79' => '', 'F79' => '', ],
		'80'   => ['B80' => '', 'C80' => '', 'E80' => '', 'F80' => '', ],
		'81'   => ['B81' => '', 'C81' => '', 'E81' => '', 'F81' => '', ],
		'82'   => ['B82' => '', 'C82' => '', 'E82' => '', 'F82' => '', ],
		'83'   => ['B83' => '', 'C83' => '', 'E83' => '', 'F83' => '', ],
		'84'   => ['B84' => '', 'C84' => '', 'E84' => '', 'F84' => '', ],
		'85'   => ['B85' => '', 'C85' => '', 'E85' => '', 'F85' => '', ],
		'86'   => ['B86' => '', 'C86' => '', 'E86' => 'I2072', 'F86' => 'U2072+AC2072', ],
		'87'   => ['B87' => '', 'C87' => '', 'E87' => 'I2073', 'F87' => 'U2073+AC2073', ],
		'88'   => ['B88' => '', 'C88' => '', 'E88' => 'I2074', 'F88' => 'U2074+AC2074', ],
		'89'   => ['B89' => '', 'C89' => '', 'E89' => 'I2075', 'F89' => 'U2075+AC2075', ],
		'90'   => ['B90' => '', 'C90' => '', 'E90' => 'I2090', 'F90' => 'U2090+AC2090', ],
		'91'   => ['B91' => '', 'C91' => '', 'E91' => 'I2091', 'F91' => 'U2091+AC2091', ],
		'92'   => ['B92' => '', 'C92' => '', 'E92' => 'I2092', 'F92' => 'U2092+AC2092', ],
		'93'   => ['B93' => '', 'C93' => '', 'E93' => 'I2108', 'F93' => 'U2108+AC2108', ],
		'94'   => ['B94' => '', 'C94' => '', 'E94' => '', 'F94' => '', ],
		'95'   => ['B95' => '', 'C95' => '', 'E95' => '', 'F95' => '', ],
		'96'   => ['B96' => '', 'C96' => '', 'E96' => '', 'F96' => '', ],
		'97'   => ['B97' => '', 'C97' => '', 'E97' => '', 'F97' => '', ],
		'98'   => ['B98' => '', 'C98' => '', 'E98' => '', 'F98' => '', ],
		'99'   => ['B99' => '', 'C99' => '', 'E99' => 'I2124', 'F99' => 'U2124+AC2124', ],
		'100'  => ['B100' => '', 'C100' => '', 'E100' => 'I2125', 'F100' => 'U2125+AC2125', ],
		'101'  => ['B101' => '', 'C101' => '', 'E101' => 'I2126', 'F101' => 'U2126+AC2126', ],
		'102'  => ['B102' => '', 'C102' => '', 'E102' => 'I2127', 'F102' => 'U2127+AC2127', ],
		'103'  => ['B103' => '', 'C103' => '', 'E103' => 'I2133', 'F103' => 'U2133+AC2133', ],
		'104'  => ['B104' => '', 'C104' => '', 'E104' => 'I2134', 'F104' => 'U2134+AC2134', ],
		'105'  => ['B105' => '', 'C105' => '', 'E105' => 'I2135', 'F105' => 'U2135+AC2135', ],
		'106'  => ['B106' => '', 'C106' => '', 'E106' => 'I1639', 'F106' => 'U1639+AC1639', ],
		'107'  => ['B107' => '', 'C107' => '', 'E107' => 'I1858', 'F107' => 'U1858+AC1858', ],
		'108'  => ['B108' => '', 'C108' => '', 'E108' => 'I1913', 'F108' => 'U1913+AC1913', ],
		'109'  => ['B109' => '', 'C109' => '', 'E109' => 'I568', 'F109' => 'U568+AC568', ],
		'110'  => ['B110' => '', 'C110' => '', 'E110' => 'I945', 'F110' => 'U945+AC945', ],
		'111'  => ['B111' => '', 'C111' => '', 'E111' => 'I946', 'F111' => 'U946+AC946', ],
		'112'  => ['B112' => '', 'C112' => '', 'E112' => 'I1107', 'F112' => 'U1107+AC1107', ],
		'113'  => ['B113' => '', 'C113' => '', 'E113' => 'I1108', 'F113' => 'U1108+AC1108', ],
		'114'  => ['B114' => '', 'C114' => '', 'E114' => 'I1109', 'F114' => 'U1109+AC1109', ],
		'115'  => ['B115' => '', 'C115' => '', 'E115' => 'I1270', 'F115' => 'U1270+AC1270', ],
		'116'  => ['B116' => '', 'C116' => '', 'E116' => 'I1273', 'F116' => 'U1273+AC1273', ],
		'117'  => ['B117' => '', 'C117' => '', 'E117' => 'I1309', 'F117' => 'U1309+AC1309', ],
		'118'  => ['B118' => '', 'C118' => '', 'E118' => 'I1406', 'F118' => 'U1406+AC1406', ],
		'119'  => ['B119' => '', 'C119' => '', 'E119' => 'I1505', 'F119' => 'U1505+AC1505', ],
		'120'  => ['B120' => '', 'C120' => '', 'E120' => 'I1514', 'F120' => 'U1514+AC1514', ],
		'121'  => ['B121' => '', 'C121' => '', 'E121' => 'I1641', 'F121' => 'U1641+AC1641', ],
		'122'  => ['B122' => '', 'C122' => '', 'E122' => 'I1787', 'F122' => 'U1787+AC1787', ],
		'123'  => ['B123' => '', 'C123' => '', 'E123' => 'I1852', 'F123' => 'U1852+AC1852', ],
		'124'  => ['B124' => '', 'C124' => '', 'E124' => 'I1879', 'F124' => 'U1879+AC1879', ],
		'125'  => ['B125' => '', 'C125' => '', 'E125' => 'I2003', 'F125' => 'U2003+AC2003', ],
		'126'  => ['B126' => '', 'C126' => '', 'E126' => 'I1231', 'F126' => 'U1231+AC1231', ],
		'127'  => ['B127' => '', 'C127' => '', 'E127' => 'I565', 'F127' => 'U565+AC565', ],
		'128'  => ['B128' => '', 'C128' => '', 'E128' => 'I900', 'F128' => 'U900+AC900', ],
		'129'  => ['B129' => '', 'C129' => '', 'E129' => 'I905', 'F129' => 'U905+AC905', ],
		'130'  => ['B130' => '', 'C130' => '', 'E130' => 'I1018', 'F130' => 'U1018+AC1018', ],
		'131'  => ['B131' => '', 'C131' => '', 'E131' => 'I1026', 'F131' => 'U1026+AC1026', ],
		'132'  => ['B132' => '', 'C132' => '', 'E132' => 'I1211', 'F132' => 'U1211+AC1211', ],
		'133'  => ['B133' => '', 'C133' => '', 'E133' => 'I1212', 'F133' => 'U1212+AC1212', ],
		'134'  => ['B134' => '', 'C134' => '', 'E134' => 'I1250', 'F134' => 'U1250+AC1250', ],
		'135'  => ['B135' => '', 'C135' => '', 'E135' => 'I1253', 'F135' => 'U1253+AC1253', ],
		'136'  => ['B136' => '', 'C136' => '', 'E136' => 'I1382', 'F136' => 'U1382+AC1382', ],
		'137'  => ['B137' => '', 'C137' => '', 'E137' => 'I1497', 'F137' => 'U1497+AC1497', ],
		'138'  => ['B138' => '', 'C138' => '', 'E138' => 'I1498', 'F138' => 'U1498+AC1498', ],
		'139'  => ['B139' => '', 'C139' => '', 'E139' => 'I1644', 'F139' => 'U1644+AC1644', ],
		'140'  => ['B140' => '', 'C140' => '', 'E140' => 'I1676', 'F140' => 'U1676+AC1676', ],
		'141'  => ['B141' => '', 'C141' => '', 'E141' => 'I1677', 'F141' => 'U1677+AC1677', ],
		'142'  => ['B142' => '', 'C142' => '', 'E142' => 'I1895', 'F142' => 'U1895+AC1895', ],
		'143'  => ['B143' => '', 'C143' => '', 'E143' => 'I1982', 'F143' => 'U1982+AC1982', ],
		'144'  => ['B144' => '', 'C144' => '', 'E144' => 'I2002', 'F144' => 'U2002+AC2002', ],
		'145'  => ['B145' => '', 'C145' => '', 'E145' => 'I524', 'F145' => 'U524+AC524', ],
		'146'  => ['B146' => '', 'C146' => '', 'E146' => 'I885', 'F146' => 'U885+AC885', ],
		'147'  => ['B147' => '', 'C147' => '', 'E147' => 'I1439', 'F147' => 'U1439+AC1439', ],
		'148'  => ['B148' => '', 'C148' => '', 'E148' => 'I1977', 'F148' => 'U1977+AC1977', ],
		'149'  => ['B149' => '', 'C149' => '', 'E149' => 'I1985', 'F149' => 'U1985+AC1985', ],
		'150'  => ['B150' => '', 'C150' => '', 'E150' => 'I2004', 'F150' => 'U2004+AC2004', ],
		'151'  => ['B151' => '', 'C151' => '', 'E151' => 'I2009', 'F151' => 'U2009+AC2009', ],
		'152'  => ['B152' => '', 'C152' => '', 'E152' => 'I2022', 'F152' => 'U2022+AC2022', ],
		'153'  => ['B153' => '', 'C153' => '', 'E153' => 'I2031', 'F153' => 'U2031+AC2031', ],
		'154'  => ['B154' => '', 'C154' => '', 'E154' => 'I2036', 'F154' => 'U2036+AC2036', ],
		'155'  => ['B155' => '', 'C155' => '', 'E155' => 'I2046', 'F155' => 'U2046+AC2046', ],
		'156'  => ['B156' => '', 'C156' => '', 'E156' => 'I2048', 'F156' => 'U2048+AC2048', ],
		'157'  => ['B157' => '', 'C157' => '', 'E157' => 'I2067', 'F157' => 'U2067+AC2067', ],
		'158'  => ['B158' => '', 'C158' => '', 'E158' => 'I2076', 'F158' => 'U2076+AC2076', ],
		'159'  => ['B159' => '', 'C159' => '', 'E159' => 'I2085', 'F159' => 'U2085+AC2085', ],
		'160'  => ['B160' => '', 'C160' => '', 'E160' => 'I2093', 'F160' => 'U2093+AC2093', ],
		'161'  => ['B161' => '', 'C161' => '', 'E161' => 'I2098', 'F161' => 'U2098+AC2098', ],
		'162'  => ['B162' => '', 'C162' => '', 'E162' => 'I2102', 'F162' => 'U2102+AC2102', ],
		'163'  => ['B163' => '', 'C163' => '', 'E163' => 'I2120', 'F163' => 'U2120+AC2120', ],
		'164'  => ['B164' => '', 'C164' => '', 'E164' => 'I2128', 'F164' => 'U2128+AC2128', ],
		'165'  => ['B165' => '', 'C165' => '', 'E165' => 'I1583', 'F165' => 'U1583+AC1583', ],
		'166'  => ['B166' => '', 'C166' => '', 'E166' => 'I258', 'F166' => 'U258+AC258', ],
		'167'  => ['B167' => '', 'C167' => '', 'E167' => 'I1810', 'F167' => 'U1810+AC1810', ],
		'168'  => ['B168' => '', 'C168' => '', 'E168' => 'I324', 'F168' => 'U324+AC324', ],
		'169'  => ['B169' => '', 'C169' => '', 'E169' => 'I326', 'F169' => 'U326+AC326', ],
		'170'  => ['B170' => '', 'C170' => '', 'E170' => 'I367', 'F170' => 'U367+AC367', ],
		'171'  => ['B171' => '', 'C171' => '', 'E171' => 'I473', 'F171' => 'U473+AC473', ],
		'172'  => ['B172' => '', 'C172' => '', 'E172' => 'I735', 'F172' => 'U735+AC735', ],
		'173'  => ['B173' => '', 'C173' => '', 'E173' => 'I841', 'F173' => 'U841+AC841', ],
		'174'  => ['B174' => '', 'C174' => '', 'E174' => 'I968', 'F174' => 'U968+AC968', ],
		'175'  => ['B175' => '', 'C175' => '', 'E175' => 'I1005', 'F175' => 'U1005+AC1005', ],
		'176'  => ['B176' => '', 'C176' => '', 'E176' => 'I1171', 'F176' => 'U1171+AC1171', ],
		'177'  => ['B177' => '', 'C177' => '', 'E177' => 'I1264', 'F177' => 'U1264+AC1264', ],
		'178'  => ['B178' => '', 'C178' => '', 'E178' => 'I1501', 'F178' => 'U1501+AC1501', ],
		'179'  => ['B179' => '', 'C179' => '', 'E179' => 'I1535', 'F179' => 'U1535+AC1535', ],
		'180'  => ['B180' => '', 'C180' => '', 'E180' => 'I1593', 'F180' => 'U1593+AC1593', ],
		'181'  => ['B181' => '', 'C181' => '', 'E181' => 'I1594', 'F181' => 'U1594+AC1594', ],
		'182'  => ['B182' => '', 'C182' => '', 'E182' => 'I1645', 'F182' => 'U1645+AC1645', ],
		'183'  => ['B183' => '', 'C183' => '', 'E183' => 'I1842', 'F183' => 'U1842+AC1842', ],
		'184'  => ['B184' => '', 'C184' => '', 'E184' => 'I1851', 'F184' => 'U1851+AC1851', ],
		'185'  => ['B185' => '', 'C185' => '', 'E185' => 'I1856', 'F185' => 'U1856+AC1856', ],
		'186'  => ['B186' => '', 'C186' => '', 'E186' => 'I899', 'F186' => 'U899+AC899', ],
		'187'  => ['B187' => '', 'C187' => '', 'E187' => 'I220', 'F187' => 'U220+AC220', ],
		'188'  => ['B188' => '', 'C188' => '', 'E188' => 'I319', 'F188' => 'U319+AC319', ],
		'189'  => ['B189' => '', 'C189' => '', 'E189' => 'I1006', 'F189' => 'U1006+AC1006', ],
		'190'  => ['B190' => '', 'C190' => '', 'E190' => 'I1265', 'F190' => 'U1265+AC1265', ],
		'191'  => ['B191' => '', 'C191' => '', 'E191' => 'I1428', 'F191' => 'U1428+AC1428', ],
		'192'  => ['B192' => '', 'C192' => '', 'E192' => 'I1516', 'F192' => 'U1516+AC1516', ],
		'193'  => ['B193' => '', 'C193' => '', 'E193' => 'I1564', 'F193' => 'U1564+AC1564', ],
		'194'  => ['B194' => '', 'C194' => '', 'E194' => 'I1679', 'F194' => 'U1679+AC1679', ],
		'195'  => ['B195' => '', 'C195' => '', 'E195' => 'I1932', 'F195' => 'U1932+AC1932', ],
		'196'  => ['B196' => '', 'C196' => '', 'E196' => 'I2023', 'F196' => 'U2023+AC2023', ],
		'197'  => ['B197' => '', 'C197' => '', 'E197' => 'I2035', 'F197' => 'U2035+AC2035', ],
		'198'  => ['B198' => '', 'C198' => '', 'E198' => 'I2068', 'F198' => 'U2068+AC2068', ],
		'199'  => ['B199' => '', 'C199' => '', 'E199' => 'I2081', 'F199' => 'U2081+AC2081', ],
		'200'  => ['B200' => '', 'C200' => '', 'E200' => 'I2094', 'F200' => 'U2094+AC2094', ],
		'201'  => ['B201' => '', 'C201' => '', 'E201' => 'I2118', 'F201' => 'U2118+AC2118', ],
		'202'  => ['B202' => '', 'C202' => '', 'E202' => 'I1425', 'F202' => 'U1425+AC1425', ],
		'203'  => ['B203' => '', 'C203' => '', 'E203' => 'I2001', 'F203' => 'U2001+AC2001', ],
		'204'  => ['B204' => '', 'C204' => '', 'E204' => 'I229', 'F204' => 'U229+AC229', ],
		'205'  => ['B205' => '', 'C205' => '', 'E205' => 'I311', 'F205' => 'U311+AC311', ],
		'206'  => ['B206' => '', 'C206' => '', 'E206' => 'I312', 'F206' => 'U312+AC312', ],
		'207'  => ['B207' => '', 'C207' => '', 'E207' => 'I313', 'F207' => 'U313+AC313', ],
		'208'  => ['B208' => '', 'C208' => '', 'E208' => 'I965', 'F208' => 'U965+AC965', ],
		'209'  => ['B209' => '', 'C209' => '', 'E209' => 'I1017', 'F209' => 'U1017+AC1017', ],
		'210'  => ['B210' => '', 'C210' => '', 'E210' => 'I1255', 'F210' => 'U1255+AC1255', ],
		'211'  => ['B211' => '', 'C211' => '', 'E211' => 'I1386', 'F211' => 'U1386+AC1386', ],
		'212'  => ['B212' => '', 'C212' => '', 'E212' => 'I1646', 'F212' => 'U1646+AC1646', ],
		'213'  => ['B213' => '', 'C213' => '', 'E213' => 'I234', 'F213' => 'U234+AC234', ],
		'214'  => ['B214' => '', 'C214' => '', 'E214' => 'I240', 'F214' => 'U240+AC240', ],
		'215'  => ['B215' => '', 'C215' => '', 'E215' => 'I321', 'F215' => 'U321+AC321', ],
		'216'  => ['B216' => '', 'C216' => '', 'E216' => 'I1057', 'F216' => 'U1057+AC1057', ],
		'217'  => ['B217' => '', 'C217' => '', 'E217' => 'I1103', 'F217' => 'U1103+AC1103', ],
		'218'  => ['B218' => '', 'C218' => '', 'E218' => 'I1104', 'F218' => 'U1104+AC1104', ],
		'219'  => ['B219' => '', 'C219' => '', 'E219' => 'I1105', 'F219' => 'U1105+AC1105', ],
		'220'  => ['B220' => '', 'C220' => '', 'E220' => 'I1106', 'F220' => 'U1106+AC1106', ],
		'221'  => ['B221' => '', 'C221' => '', 'E221' => 'I1474', 'F221' => 'U1474+AC1474', ],
		'222'  => ['B222' => '', 'C222' => '', 'E222' => 'I228', 'F222' => 'U228+AC228', ],
		'223'  => ['B223' => '', 'C223' => '', 'E223' => 'I2103', 'F223' => 'U2103+AC2103', ],
		'224'  => ['B224' => '', 'C224' => '', 'E224' => 'I2010', 'F224' => 'U2010+AC2010', ],
		'225'  => ['B225' => '', 'C225' => '', 'E225' => 'I238', 'F225' => 'U238+AC238', ],
		'226'  => ['B226' => '', 'C226' => '', 'E226' => 'I245', 'F226' => 'U245+AC245', ],
		'227'  => ['B227' => '', 'C227' => '', 'E227' => 'I371', 'F227' => 'U371+AC371', ],
		'228'  => ['B228' => '', 'C228' => '', 'E228' => 'I846', 'F228' => 'U846+AC846', ],
		'229'  => ['B229' => '', 'C229' => '', 'E229' => 'I1266', 'F229' => 'U1266+AC1266', ],
		'230'  => ['B230' => '', 'C230' => '', 'E230' => 'I1267', 'F230' => 'U1267+AC1267', ],
		'231'  => ['B231' => '', 'C231' => '', 'E231' => 'I1479', 'F231' => 'U1479+AC1479', ],
		'232'  => ['B232' => '', 'C232' => '', 'E232' => 'I1536', 'F232' => 'U1536+AC1536', ],
		'233'  => ['B233' => '', 'C233' => '', 'E233' => 'I2070', 'F233' => 'U2070+AC2070', ],
		'234'  => ['B234' => '', 'C234' => '', 'E234' => 'I360', 'F234' => 'U360+AC360', ],
		'235'  => ['B235' => '', 'C235' => '', 'E235' => 'I362', 'F235' => 'U362+AC362', ],
		'236'  => ['B236' => '', 'C236' => '', 'E236' => 'I462', 'F236' => 'U462+AC462', ],
		'237'  => ['B237' => '', 'C237' => '', 'E237' => 'I853', 'F237' => 'U853+AC853', ],
		'238'  => ['B238' => '', 'C238' => '', 'E238' => 'I1007', 'F238' => 'U1007+AC1007', ],
		'239'  => ['B239' => '', 'C239' => '', 'E239' => 'I1009', 'F239' => 'U1009+AC1009', ],
		'240'  => ['B240' => '', 'C240' => '', 'E240' => 'I1234', 'F240' => 'U1234+AC1234', ],
		'241'  => ['B241' => '', 'C241' => '', 'E241' => 'I1299', 'F241' => 'U1299+AC1299', ],
		'242'  => ['B242' => '', 'C242' => '', 'E242' => 'I1312', 'F242' => 'U1312+AC1312', ],
		'243'  => ['B243' => '', 'C243' => '', 'E243' => 'I1313', 'F243' => 'U1313+AC1313', ],
		'244'  => ['B244' => '', 'C244' => '', 'E244' => 'I1380', 'F244' => 'U1380+AC1380', ],
		'245'  => ['B245' => '', 'C245' => '', 'E245' => 'I1381', 'F245' => 'U1381+AC1381', ],
		'246'  => ['B246' => '', 'C246' => '', 'E246' => 'I1414', 'F246' => 'U1414+AC1414', ],
		'247'  => ['B247' => '', 'C247' => '', 'E247' => 'I1430', 'F247' => 'U1430+AC1430', ],
		'248'  => ['B248' => '', 'C248' => '', 'E248' => 'I1649', 'F248' => 'U1649+AC1649', ],
		'249'  => ['B249' => '', 'C249' => '', 'E249' => 'I1774', 'F249' => 'U1774+AC1774', ],
		'250'  => ['B250' => '', 'C250' => '', 'E250' => 'I1775', 'F250' => 'U1775+AC1775', ],
		'251'  => ['B251' => '', 'C251' => '', 'E251' => 'I1980', 'F251' => 'U1980+AC1980', ],
		'252'  => ['B252' => '', 'C252' => '', 'E252' => 'I1988', 'F252' => 'U1988+AC1988', ],
		'253'  => ['B253' => '', 'C253' => '', 'E253' => 'I2007', 'F253' => 'U2007+AC2007', ],
		'254'  => ['B254' => '', 'C254' => '', 'E254' => 'I2025', 'F254' => 'U2025+AC2025', ],
		'255'  => ['B255' => '', 'C255' => '', 'E255' => 'I2033', 'F255' => 'U2033+AC2033', ],
		'256'  => ['B256' => '', 'C256' => '', 'E256' => 'I2039', 'F256' => 'U2039+AC2039', ],
		'257'  => ['B257' => '', 'C257' => '', 'E257' => 'I2047', 'F257' => 'U2047+AC2047', ],
		'258'  => ['B258' => '', 'C258' => '', 'E258' => 'I2050', 'F258' => 'U2050+AC2050', ],
		'259'  => ['B259' => '', 'C259' => '', 'E259' => 'I2079', 'F259' => 'U2079+AC2079', ],
		'260'  => ['B260' => '', 'C260' => '', 'E260' => 'I2080', 'F260' => 'U2087+AC2087', ],
		'261'  => ['B261' => '', 'C261' => '', 'E261' => 'I2081', 'F261' => 'U2096+AC2096', ],
		'262'  => ['B262' => '', 'C262' => '', 'E262' => 'I2082', 'F262' => 'U2101+AC2101', ],
		'263'  => ['B263' => '', 'C263' => '', 'E263' => 'I2083', 'F263' => 'U2105+AC2105', ],
		'264'  => ['B264' => '', 'C264' => '', 'E264' => 'I2123', 'F264' => 'U2123+AC2123', ],
		'265'  => ['B265' => '', 'C265' => '', 'E265' => 'I1979', 'F265' => 'U1979+AC1979', ],
		'266'  => ['B266' => '', 'C266' => '', 'E266' => 'I1987', 'F266' => 'U1987+AC1987', ],
		'267'  => ['B267' => '', 'C267' => '', 'E267' => 'I2006', 'F267' => 'U2006+AC2006', ],
		'268'  => ['B268' => '', 'C268' => '', 'E268' => 'I2011', 'F268' => 'U2011+AC2011', ],
		'269'  => ['B269' => '', 'C269' => '', 'E269' => 'I2024', 'F269' => 'U2024+AC2024', ],
		'270'  => ['B270' => '', 'C270' => '', 'E270' => 'I2032', 'F270' => 'U2032+AC2032', ],
		'271'  => ['B271' => '', 'C271' => '', 'E271' => 'I2038', 'F271' => 'U2038+AC2038', ],
		'272'  => ['B272' => '', 'C272' => '', 'E272' => 'I2042', 'F272' => 'U2042+AC2042', ],
		'273'  => ['B273' => '', 'C273' => '', 'E273' => 'I2044', 'F273' => 'U2044+AC2044', ],
		'274'  => ['B274' => '', 'C274' => '', 'E274' => 'I2078', 'F274' => 'U2078+AC2078', ],
		'275'  => ['B275' => '', 'C275' => '', 'E275' => 'I2084', 'F275' => 'U2084+AC2084', ],
		'276'  => ['B276' => '', 'C276' => '', 'E276' => 'I2095', 'F276' => 'U2095+AC2095', ],
		'277'  => ['B277' => '', 'C277' => '', 'E277' => 'I2100', 'F277' => 'U2100+AC2100', ],
		'278'  => ['B278' => '', 'C278' => '', 'E278' => 'I2104', 'F278' => 'U2104+AC2104', ],
		'279'  => ['B279' => '', 'C279' => '', 'E279' => 'I2105', 'F279' => 'U2122+AC2122', ],
		'280'  => ['B280' => '', 'C280' => '', 'E280' => 'I2130', 'F280' => 'U2130+AC2130', ],
		'281'  => ['B281' => '', 'C281' => '', 'E281' => 'I2069', 'F281' => 'U2069+AC2069', ],
		'282'  => ['B282' => '', 'C282' => '', 'E282' => 'I1984', 'F282' => 'U1984+AC1984', ],
		'283'  => ['B283' => '', 'C283' => '', 'E283' => 'I2131', 'F283' => 'U2131+AC2131', ],
		'284'  => ['B284' => '', 'C284' => '', 'E284' => 'I2130', 'F284' => 'U2130+AC2130', ],
		'285'  => ['B285' => '', 'C285' => '', 'E285' => 'I1989', 'F285' => 'U1989+AC1989', ],
		'286'  => ['B286' => '', 'C286' => '', 'E286' => 'I2008', 'F286' => 'U2008+AC2008', ],
		'287'  => ['B287' => '', 'C287' => '', 'E287' => 'I2013', 'F287' => 'U2013+AC2013', ],
		'288'  => ['B288' => '', 'C288' => '', 'E288' => 'I2040', 'F288' => 'U2040+AC2040', ],
		'289'  => ['B289' => '', 'C289' => '', 'E289' => 'I2043', 'F289' => 'U2043+AC2043', ],
		'290'  => ['B290' => '', 'C290' => '', 'E290' => 'I2071', 'F290' => 'U2071+AC2071', ],
		'291'  => ['B291' => '', 'C291' => '', 'E291' => 'I2097', 'F291' => 'U2097+AC2097', ],
		'292'  => ['B292' => '', 'C292' => '', 'E292' => 'I2106', 'F292' => 'U2106+AC2106', ],
		'293'  => ['B293' => '', 'C293' => '', 'E293' => 'I2132', 'F293' => 'U2132+AC2132', ],
		'294'  => ['B294' => '', 'C294' => '', 'E294' => 'I1118', 'F294' => 'U1118+AC1118', ],
		'295'  => ['B295' => '', 'C295' => '', 'E295' => 'I1680', 'F295' => 'U1680+AC1680', ],
		'296'  => ['B296' => '', 'C296' => '', 'E296' => 'I1687', 'F296' => 'U1736+AC1736', ],
		'297'  => ['B297' => '', 'C297' => '', 'E297' => 'I1894', 'F297' => 'U1894+AC1894', ],
		'298'  => ['B298' => '', 'C298' => '', 'E298' => 'I1508', 'F298' => 'U1508+AC1508', ],
		'299'  => ['B299' => '', 'C299' => '', 'E299' => 'I1683', 'F299' => 'U1683+AC1683', ],
		'300'  => ['B300' => '', 'C300' => '', 'E300' => 'I1684', 'F300' => 'U1684+AC1684', ],
		'301'  => ['B301' => '', 'C301' => '', 'E301' => 'I274', 'F301' => 'U274+AC274', ],
		'302'  => ['B302' => '', 'C302' => '', 'E302' => 'I275', 'F302' => 'U275+AC275', ],
		'303'  => ['B303' => '', 'C303' => '', 'E303' => 'I421', 'F303' => 'U421+AC421', ],
		'304'  => ['B304' => '', 'C304' => '', 'E304' => 'I422', 'F304' => 'U422+AC422', ],
		'305'  => ['B305' => '', 'C305' => '', 'E305' => 'I423', 'F305' => 'U423+AC423', ],
		'306'  => ['B306' => '', 'C306' => '', 'E306' => 'I446', 'F306' => 'U446+AC446', ],
		'307'  => ['B307' => '', 'C307' => '', 'E307' => 'I484', 'F307' => 'U484+AC484', ],
		'308'  => ['B308' => '', 'C308' => '', 'E308' => 'I922', 'F308' => 'U922+AC922', ],
		'309'  => ['B309' => '', 'C309' => '', 'E309' => 'I951', 'F309' => 'U951+AC951', ],
		'310'  => ['B310' => '', 'C310' => '', 'E310' => 'I956', 'F310' => 'U956+AC956', ],
		'311'  => ['B311' => '', 'C311' => '', 'E311' => 'I1277', 'F311' => 'U1277+AC1277', ],
		'312'  => ['B312' => '', 'C312' => '', 'E312' => 'I1281', 'F312' => 'U1281+AC1281', ],
		'313'  => ['B313' => '', 'C313' => '', 'E313' => 'I1282', 'F313' => 'U1282+AC1282', ],
		'314'  => ['B314' => '', 'C314' => '', 'E314' => 'I1642', 'F314' => 'U1642+AC1642', ],
		'315'  => ['B315' => '', 'C315' => '', 'E315' => 'I1938', 'F315' => 'U1938+AC1938', ],
		'316'  => ['B316' => '', 'C316' => '', 'E316' => 'I373', 'F316' => 'U373+AC373', ],
		'317'  => ['B317' => '', 'C317' => '', 'E317' => 'I374', 'F317' => 'U374+AC374', ],
		'318'  => ['B318' => '', 'C318' => '', 'E318' => 'I375', 'F318' => 'U375+AC375', ],
		'319'  => ['B319' => '', 'C319' => '', 'E319' => 'I376', 'F319' => 'U376+AC376', ],
		'320'  => ['B320' => '', 'C320' => '', 'E320' => 'I924', 'F320' => 'U924+AC924', ],
		'321'  => ['B321' => '', 'C321' => '', 'E321' => 'I954', 'F321' => 'U954+AC954', ],
		'322'  => ['B322' => '', 'C322' => '', 'E322' => 'I955', 'F322' => 'U957+AC957', ],
		'323'  => ['B323' => '', 'C323' => '', 'E323' => 'I1092', 'F323' => 'U1092+AC1092', ],
		'324'  => ['B324' => '', 'C324' => '', 'E324' => 'I1093', 'F324' => 'U1093+AC1093', ],
		'325'  => ['B325' => '', 'C325' => '', 'E325' => 'I1094', 'F325' => 'U1094+AC1094', ],
		'326'  => ['B326' => '', 'C326' => '', 'E326' => 'I1096', 'F326' => 'U1096+AC1096', ],
		'327'  => ['B327' => '', 'C327' => '', 'E327' => 'I1097', 'F327' => 'U1097+AC1097', ],
		'328'  => ['B328' => '', 'C328' => '', 'E328' => 'I1098', 'F328' => 'U1098+AC1098', ],
		'329'  => ['B329' => '', 'C329' => '', 'E329' => 'I1175', 'F329' => 'U1175+AC1175', ],
		'330'  => ['B330' => '', 'C330' => '', 'E330' => 'I1279', 'F330' => 'U1279+AC1279', ],
		'331'  => ['B331' => '', 'C331' => '', 'E331' => 'I1392', 'F331' => 'U1392+AC1392', ],
		'332'  => ['B332' => '', 'C332' => '', 'E332' => 'I1402', 'F332' => 'U1402+AC1402', ],
		'333'  => ['B333' => '', 'C333' => '', 'E333' => 'I1456', 'F333' => 'U1456+AC1456', ],
		'334'  => ['B334' => '', 'C334' => '', 'E334' => 'I1457', 'F334' => 'U1457+AC1457', ],
		'335'  => ['B335' => '', 'C335' => '', 'E335' => 'I1648', 'F335' => 'U1648+AC1648', ],
		'336'  => ['B336' => '', 'C336' => '', 'E336' => 'I2012', 'F336' => 'U2012+AC2012', ],
		'337'  => ['B337' => '', 'C337' => '', 'E337' => 'I1638', 'F337' => 'U1638+AC1638', ],
		'338'  => ['B338' => '', 'C338' => '', 'E338' => 'I318', 'F338' => 'U318+AC318', ],
		'339'  => ['B339' => '', 'C339' => '', 'E339' => 'I377', 'F339' => 'U377+AC377', ],
		'340'  => ['B340' => '', 'C340' => '', 'E340' => 'I379', 'F340' => 'U379+AC379', ],
		'341'  => ['B341' => '', 'C341' => '', 'E341' => 'I472', 'F341' => 'U472+AC472', ],
		'342'  => ['B342' => '', 'C342' => '', 'E342' => 'I500', 'F342' => 'U500+AC500', ],
		'343'  => ['B343' => '', 'C343' => '', 'E343' => 'I502', 'F343' => 'U502+AC502', ],
		'344'  => ['B344' => '', 'C344' => '', 'E344' => 'I503', 'F344' => 'U503+AC503', ],
		'345'  => ['B345' => '', 'C345' => '', 'E345' => 'I504', 'F345' => 'U504+AC504', ],
		'346'  => ['B346' => '', 'C346' => '', 'E346' => 'I505', 'F346' => 'U506+AC506', ],
		'347'  => ['B347' => '', 'C347' => '', 'E347' => 'I506', 'F347' => 'U507+AC507', ],
		'348'  => ['B348' => '', 'C348' => '', 'E348' => 'I518', 'F348' => 'U518+AC518', ],
		'349'  => ['B349' => '', 'C349' => '', 'E349' => 'I519', 'F349' => 'U519+AC519', ],
		'350'  => ['B350' => '', 'C350' => '', 'E350' => 'I909', 'F350' => 'U909+AC909', ],
		'351'  => ['B351' => '', 'C351' => '', 'E351' => 'I910', 'F351' => 'U926+AC926', ],
		'352'  => ['B352' => '', 'C352' => '', 'E352' => 'I948', 'F352' => 'U948+AC948', ],
		'353'  => ['B353' => '', 'C353' => '', 'E353' => 'I1095', 'F353' => 'U1095+AC1095', ],
		'354'  => ['B354' => '', 'C354' => '', 'E354' => 'I1115', 'F354' => 'U1115+AC1115', ],
		'355'  => ['B355' => '', 'C355' => '', 'E355' => 'I1232', 'F355' => 'U1232+AC1232', ],
		'356'  => ['B356' => '', 'C356' => '', 'E356' => 'I1283', 'F356' => 'U1283+AC1283', ],
		'357'  => ['B357' => '', 'C357' => '', 'E357' => 'I1285', 'F357' => 'U1285+AC1285', ],
		'358'  => ['B358' => '', 'C358' => '', 'E358' => 'I1286', 'F358' => 'U1286+AC1286', ],
		'359'  => ['B359' => '', 'C359' => '', 'E359' => 'I1288', 'F359' => 'U1288+AC1288', ],
		'360'  => ['B360' => '', 'C360' => '', 'E360' => 'I1288', 'F360' => 'U1298+AC1298', ],
		'361'  => ['B361' => '', 'C361' => '', 'E361' => 'I1314', 'F361' => 'U1314+AC1314', ],
		'362'  => ['B362' => '', 'C362' => '', 'E362' => 'I1455', 'F362' => 'U1455+AC1455', ],
		'363'  => ['B363' => '', 'C363' => '', 'E363' => 'I1489', 'F363' => 'U1489+AC1489', ],
		'364'  => ['B364' => '', 'C364' => '', 'E364' => 'I1490', 'F364' => 'U1490+AC1490', ],
		'365'  => ['B365' => '', 'C365' => '', 'E365' => 'I1500', 'F365' => 'U1500+AC1500', ],
		'366'  => ['B366' => '', 'C366' => '', 'E366' => 'I1517', 'F366' => 'U1517+AC1517', ],
		'367'  => ['B367' => '', 'C367' => '', 'E367' => 'I1518', 'F367' => 'U1529+AC1529', ],
		'368'  => ['B368' => '', 'C368' => '', 'E368' => 'I1681', 'F368' => 'U1681+AC1681', ],
		'369'  => ['B369' => '', 'C369' => '', 'E369' => 'I1682', 'F369' => 'U1682+AC1682', ],
		'370'  => ['B370' => '', 'C370' => '', 'E370' => 'I1685', 'F370' => 'U1685+AC1685', ],
		'371'  => ['B371' => '', 'C371' => '', 'E371' => 'I1792', 'F371' => 'U1792+AC1792', ],
		'372'  => ['B372' => '', 'C372' => '', 'E372' => 'I1710', 'F372' => 'U1710+AC1710', ],
		'373'  => ['B373' => '', 'C373' => '', 'E373' => 'I1937', 'F373' => 'U1937+AC1937', ],
		'374'  => ['B374' => '', 'C374' => '', 'E374' => 'I1803', 'F374' => 'U1803+AC1803', ],
		'375'  => ['B375' => '$J1317', 'C375' => '$V1317+$AD1317', 'E375' => 'I1317', 'F375' => 'U1317+AC1317', ],
		'376'  => ['B376' => '$I2296+$J2296-SUM($B8:$B19,$B377,$B378)', 'C376' => 'SUM($U2296:$V2296,$AC2294:$AD2294)-SUM($C8:$C19,$C377,$C378)', 'E376' => '$I2296-SUM($E8:$E375,$E377,$E378)+I1', 'F376' => 'SUM($U2296,$AC2294)-SUM($F8:$F375,$F377,$F378)', ],
		'377'  => ['B377' => 'SUM($I186:$J186)', 'C377' => '$U186+$V186+$AC186+$AD186', 'E377' => 'SUM($I186:$I186)', 'F377' => '$U186+$AC186', ],
		'378'  => ['B378' => 'SUM($I1970:$J1970)', 'C378' => '$U1970+$V1970+$AC1970+$AD1970', 'E378' => 'SUM($I1970:$I1970)', 'F378' => '$U1970+$AC1970', ],
		'379'  => ['B379' => 'SUM($B8:$B19,$B377:$B378)', 'C379' => 'SUM($C8:$C19,$C377:$C378)', 'E379' => 'SUM($E8:$E374,$E377:$E378)', 'F379' => 'SUM($F8:$F374,$F377:$F378)', ],
		'380'  => ['B380' => '$Q2296+$B376', 'C380' => '$Y2296+$AG2294+$C376', 'E380' => '$Q2296+$J2296+$E376', 'F380' => '$Y2296+$V2296+$AG2294+$AD2294+$F376', ],
		'381'  => ['B381' => 'SUM($B379,$B380)+I1', 'C381' => 'SUM($C379,$C380)', 'E381' => 'SUM(E379,E380)', 'F381' => 'SUM($F379,$F380)', ],
		'385'  => ['B385' => 'J2197', 'C385' => 'V2197+AD2197', ],
		'386'  => ['B386' => 'J114', 'C386' => 'V114+AD114', ],
		'387'  => ['E387' => '', 'F387' => '', ],
		'388'  => ['B388' => 'J116', 'C388' => 'V116+AD116', ],
	];


$colIndexes = [];




function showFormulas($arr)
{
	$result = "[\n";
	foreach ($arr as $k => $subArr) {
		$result .= "\t" . Filter::strPad("'$k'", 6) . " => [";
		foreach ($subArr as $key => $cell) {
			$result .= $result ? ', ' : '';
			$result .= "'$key' => '$cell'";
		}
		$result .= "],\n";
	}
	return "$result\n]";
}






/**
 * Вытяжка из куска xlsx нужных для отчёта формул.
 * В принципе, можно и в проект это встроить, и автоматизировать, открывая xlsx как zip, но надо ли?
 */
function getFormulas()
{
	$filename = CONFIG::ROOT . '/data/sheet2.xml';
	$xml = $xml = simplexml_load_file($filename);

	$rowCounter = 1;
	$result = [];
	foreach ($xml->sheetData->row as $row) {
		if ($rowCounter > 7) {
			$counter = 0;
			foreach ($row as $c) {
				if ($counter > 6) {
					break;
				}
				$counter++;
				$attr = $c->attributes();
				$ch = strval($attr['r'][0]);
				if (in_array($ch{0}, ['B', 'C', 'E', 'F'])) {
					$result[$rowCounter][$ch] =
						isset($c->f) ? (string)$c->f : '';
				}
			}
		}
		$rowCounter++;
		if ($rowCounter > 390) {
			break;
		}
	}

	return $result;
}