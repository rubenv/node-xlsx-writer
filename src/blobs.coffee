module.exports =
    contentTypes: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        </Types>
    """.replace(/\n\s*/g, '')

    rels: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
    """.replace(/\n\s*/g, '')

    workbook: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
            <workbookPr defaultThemeVersion="124226"/>
            <bookViews>
                <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
            </bookViews>
            <sheets>
                <sheet name="hello" sheetId="1" r:id="rId1"/>
            </sheets>
            <calcPr calcId="145621"/>
        </workbook>
    """.replace(/\n\s*/g, '')

    workbookRels: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
    """.replace(/\n\s*/g, '')

    styles: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            <fonts count="1" x14ac:knownFonts="1">
                <font>
                    <sz val="11"/>
                    <color theme="1"/>
                    <name val="Calibri"/>
                    <family val="2"/>
                    <scheme val="minor"/>
                </font>
            </fonts>
            <fills count="2">
                <fill>
                    <patternFill patternType="none"/>
                </fill>
                <fill>
                    <patternFill patternType="gray125"/>
                </fill>
            </fills>
            <borders count="1">
                <border>
                    <left/>
                    <right/>
                    <top/>
                    <bottom/>
                    <diagonal/>
                </border>
            </borders>
            <cellStyleXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            </cellStyleXfs>
            <cellXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            </cellXfs>
            <cellStyles count="1">
                <cellStyle name="Normal" xfId="0" builtinId="0"/>
            </cellStyles>
            <dxfs count="0"/>
            <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
            <extLst>
                <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                    <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
                </ext>
            </extLst>
        </styleSheet>
    """.replace(/\n\s*/g, '')

    strings: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
            <si>
                <t>Ruben</t>
            </si>
        </sst>
    """.replace(/\n\s*/g, '')

    sheet: """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            <dimension ref="A1:B1"/>
            <sheetViews>
                <sheetView workbookViewId="0"/>
            </sheetViews>
            <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
            <sheetData>
                <row r="1">
                    <c r="A1" t="n">
                        <v>100</v>
                    </c>
                    <c r="B1" t="s">
                        <v>0</v>
                    </c>
                </row>
            </sheetData>
        </worksheet>
    """.replace(/\n\s*/g, '')

