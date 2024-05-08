using PPTX
using Test
using DataFrames
using EzXML
using ZipArchives: ZipBufferReader, zip_readentry

@testset "PPTX Tables from a DataFrame" begin
    df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
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

@testset "check empty table style list" begin
    tableStyles_path = abspath(joinpath(artifact"pptx_data", "templates", "tableStyles.xml"))
    table_style_doc = EzXML.parsexml(read(tableStyles_path))
    @test !PPTX.has_empty_table_list(table_style_doc)

    no_slides_template = ZipBufferReader(read(joinpath(artifact"pptx_data", "templates", "no-slides.pptx")))
    table_style_doc = EzXML.parsexml(zip_readentry(no_slides_template, "ppt/tableStyles.xml"))
    @test PPTX.has_empty_table_list(table_style_doc)
end