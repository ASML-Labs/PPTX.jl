using PPTX
using Test
using DataFrames
using EzXML
using ZipArchives: ZipBufferReader, zip_readentry
using Colors

@testset "PPTX Tables from a Matrix" begin
    t = Table(rand(3,2))
    @test t.header == false

    @test PPTX.ncols(t) == 2
    @test PPTX.nrows(t) == 3

    @test_throws AssertionError Table(rand(3,2); column_widths = [20,])
    @test_throws AssertionError Table(rand(3,2); row_heights = [20,])
end

@testset "PPTX Tables from a DataFrame" begin
    lines = PPTX.TableLines(left=(width=1,))
    @test lines.left.width == 12700 # EMUs

    @test_throws AssertionError PPTX.TableLines(left=(width=1,dash=:dashss))

    nt = (left=(width=1,), top=(color=:green,))
    lines = PPTX.TableLines(nt)
    @test lines.left.width == 12700 # EMUs
    @test lines.top.color == hex(colorant"green")

    t_margins = TableCell(3; margins=(bottom=0.1,))
    @test PPTX.has_margins(t_margins)
    @test t_margins.margins.bottom == 36000
    @test t_margins.margins.left === nothing

    @test TableCell(4; color = :green).color == hex(colorant"green")

    t4 = TableCell(4; color = :green, lines = nt)
    @test lines.top.color == hex(colorant"green")
    @test PPTX.has_tc_properties(t4)
    @test !PPTX.has_margins(t4)

    t3 = TableCell(
        3;
        lines=(bottom=(width=3,color=:black),),
        style=(align=:center,),
        anchor=:center,
        color=missing,
        margins=(bottom=0.1,)
    )
    @test PPTX.has_tc_properties(t3)

    df = DataFrame(a = [1,TableCell(2)], b = [3,t4], c = [5,6])

    t = Table(df; header=false)
    @test t.header == false

    t = Table(df; offset_x=50, offset_y=50, size_x=200, size_y=150)
    contains(PPTX._show_string(t, false), "content isa DataFrame")

    @test t.content === df
    @test t.size_x == 200*PPTX._EMUS_PER_MM

    @testset "make_tblPr to XML" begin
        xml_dict = PPTX.make_tblPr(t)
        doc = PPTX.xml_document(xml_dict)
        firstRow = findfirst("//@firstRow", root(doc))
        @test firstRow.type == EzXML.ATTRIBUTE_NODE
        @test firstRow.content == "1"
        tableStyleId = findfirst("//*[local-name()='a:tableStyleId']", root(doc))
        @test tableStyleId.type == EzXML.ELEMENT_NODE
        @test tableStyleId.content == t.style_id
    end

    @testset "make_tblGrid" begin
        xml_dict = PPTX.make_tblGrid(t)
        nr_cols = size(t.content)[2]
        @test length(xml_dict["a:tblGrid"]) == nr_cols
        @test xml_dict["a:tblGrid"][1]["a:gridCol"][1]["w"] == "$(t.size_x÷nr_cols)"

        t2 = Table(df[:,1:2])
        xml_dict = PPTX.make_tblGrid(t2)
        nr_cols = size(t2.content)[2]
        @test length(xml_dict["a:tblGrid"]) == nr_cols
        @test xml_dict["a:tblGrid"][1]["a:gridCol"][1]["w"] == "$(t2.size_x÷nr_cols)"
    end

    @testset "XML dict of Table" begin
        xml_dict = PPTX.make_xml(t, 5)
        @test haskey(xml_dict, "p:graphicFrame")
        tbl = xml_dict["p:graphicFrame"][end]["a:graphic"]["a:graphicData"][end]
        @test length(tbl)==1 && haskey(tbl, "a:tbl")
        doc = PPTX.xml_document(tbl)
        elements = findall("//*[local-name()='a:tc']", root(doc))
        sz = size(t.content)
        nr_columnnames = sz[2]
        nr_elements = sz[1]*sz[2] + nr_columnnames
        @test length(elements) == nr_elements
    end
end

@testset "TableCell anchor" begin
    t = TableCell(1)
    @test PPTX.make_anchor(t) === nothing

    t = TableCell(1; anchor="center")
    @test PPTX.make_anchor(t)["anchor"] == "ctr"

    t = TableCell(1; anchor="top")
    @test PPTX.make_anchor(t)["anchor"] == "t"

    t = TableCell(1; anchor="bottom")
    @test PPTX.make_anchor(t)["anchor"] == "b"

    msg = "unknown table cell anchor \"bla\""
    @test_throws AssertionError TableCell(1; anchor="bla")
end

@testset "TableCell direction" begin

end

@testset "check empty table style list" begin
    tableStyles_path = abspath(joinpath(PPTX.TEMPLATE_DIR, "tableStyles.xml"))
    table_style_doc = EzXML.parsexml(read(tableStyles_path))
    @test !PPTX.has_empty_table_list(table_style_doc)

    no_slides_template = ZipBufferReader(read(joinpath(PPTX.TEMPLATE_DIR, "no-slides.pptx")))
    table_style_doc = EzXML.parsexml(zip_readentry(no_slides_template, "ppt/tableStyles.xml"))
    @test PPTX.has_empty_table_list(table_style_doc)
end