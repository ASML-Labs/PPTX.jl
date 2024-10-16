using Test
using PPTX
using EzXML
using ZipArchives: ZipBufferReader, zip_readentry

@testset "Slide XML structure" begin
    text_box = TextBox(content="bla", style = TextStyle(bold = true, italic = true, fontsize = 24))
    style_xml = PPTX.text_style_xml(text_box.content)
    @test any(x->get(x, "i", false)=="1", style_xml)
    @test any(x->get(x, "b", false)=="1", style_xml)
    @test any(x->get(x, "sz", false)=="2400", style_xml)

    s = Slide()
    text_box = TextBox(content="bla")

    push!(s, text_box)
    xml = PPTX.make_slide(s)
    @test haskey(xml, "p:sld")

    sld = xml["p:sld"]
    @test issubset(PPTX.main_attributes(), sld)
    csld_index = findfirst(x -> haskey(x, "p:cSld"), sld)
    @test length(sld[csld_index]["p:cSld"]) == 1
    @test haskey(sld[csld_index]["p:cSld"][begin], "p:spTree")

    # this is where the slide shapes are stored
    spTree = sld[csld_index]["p:cSld"][begin]["p:spTree"]
end

@testset "Slide Relationships XML structure" begin
    s = Slide(;layout=1)
    xml = PPTX.make_slide_relationships(s)
    # currently hardcoded that 2nd element is the layout
    layout_relationship = xml["Relationships"][2]["Relationship"]
    @test layout_relationship["Target"] == "../slideLayouts/slideLayout1.xml"

    s.layout = 2
    xml = PPTX.make_slide_relationships(s)
    layout_relationship = xml["Relationships"][2]["Relationship"]
    @test layout_relationship["Target"] == "../slideLayouts/slideLayout2.xml"
end

@testset "rId always bigger than 1 updating on push!" begin
    s = Slide()
    push!(s, TextBox("Some text"))
    @test PPTX.rid(s.shapes[1]) == 0 # textboxes have no rid, so set to 0
    push!(s, Picture(joinpath(PPTX.ASSETS_DIR,"julia_logo.png")))
    @test PPTX.rid(s.shapes[2]) == 2
    push!(s, Picture(joinpath(PPTX.ASSETS_DIR,"cauliflower.jpg")))
    @test PPTX.rid(s.shapes[3]) == 3
end

@testset "update title in XML" begin
    template = ZipBufferReader(read(joinpath(PPTX.TEMPLATE_DIR,"no-slides.pptx")))

    @testset "slideLayout1.xml" begin
        slide = Slide(;layout=1)
        layout_path = "ppt/slideLayouts/slideLayout$(slide.layout).xml"
        sp_node = PPTX.get_title_shape_node(EzXML.parsexml(zip_readentry(template, layout_path)))

        # check we can mutate the title
        t = PPTX.get_title_from_shape_node(sp_node)
        @test t.content == "Click to edit Master title style"
        PPTX.update_xml_title!(sp_node, "my own title")
        t = PPTX.get_title_from_shape_node(sp_node)
        @test t.content == "my own title"

        # check we can mutate the id
        cNvPr = findfirst("./p:nvSpPr/p:cNvPr", sp_node)
        @test cNvPr["id"] != "4"
        PPTX.update_shape_id!(sp_node, 4)
        cNvPr = findfirst("./p:nvSpPr/p:cNvPr", sp_node)
        @test cNvPr["id"] == "4"
    end

    @testset "slideLayout2.xml" begin
        slide = Slide(;layout=2)

        layout_path = "ppt/slideLayouts/slideLayout$(slide.layout).xml"
        sp_node = PPTX.get_title_shape_node(EzXML.parsexml(zip_readentry(template, layout_path)))
        @test !isnothing(sp_node)
        t = PPTX.get_title_from_shape_node(sp_node)
        @test t.content == "Click to edit Master title style"

        xml = PPTX.make_slide(slide)
        doc = PPTX.xml_document(xml)
        PPTX.add_title_shape!(doc, slide, template)
    end

    @testset "test unique shape ids" begin
        slide = Slide(;layout=2, title="my title")
        text_box = TextBox(content="bla")
        push!(slide, text_box)
        # testing all properties of the TextBox
        text_box = TextBox(content="content", offset_x=100, offset_y=140, style = Dict("bold" => true, "italic" => true))
        push!(slide, text_box)
        xml = PPTX.make_slide(slide)
        doc = PPTX.xml_document(xml)
        ids = PPTX.get_shape_ids(doc)
        @test ids == [1,2,3]

        PPTX.add_title_shape!(doc, slide, template)
        ids = PPTX.get_shape_ids(doc)
        @test ids == [1,2,3,4]
    end
end

