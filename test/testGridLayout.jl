@testset "GridLayout" begin
    @testset "constructor" begin
        layout = GridLayout(2, 3)
        @test layout.nrows == 2
        @test layout.ncols == 3
        @test layout.padding == PPTX.mm_to_emu(5)
        @test layout.margins == PPTX.GridMargins(5)
        @test isempty(layout._entries)

        layout_p = GridLayout(3, 3; padding=10)
        @test layout_p.padding == PPTX.mm_to_emu(10)
        @test layout_p.margins == PPTX.GridMargins(10)

        @test_throws ArgumentError GridLayout(0, 2)
        @test_throws ArgumentError GridLayout(2, 0)
        @test_throws ArgumentError GridLayout(2, 2; padding=-1)
        @test_throws ArgumentError GridLayout(2, 2; margins=(left=-1, right=0, top=0, bottom=0))
    end

    @testset "setindex! integer indices" begin
        layout = GridLayout(2, 2)
        box = TextBox("Hello")
        layout[1, 1] = box
        @test length(layout._entries) == 1
        shape, rows, cols = layout._entries[1]
        @test shape === box
        @test rows == 1:1
        @test cols == 1:1
    end

    @testset "setindex! colon spans" begin
        layout = GridLayout(2, 3)
        box = TextBox("Wide")
        layout[2, :] = box
        shape, rows, cols = layout._entries[1]
        @test rows == 2:2
        @test cols == 1:3

        layout2 = GridLayout(3, 2)
        layout2[:, 1] = TextBox("Tall")
        _, rows2, cols2 = layout2._entries[1]
        @test rows2 == 1:3
        @test cols2 == 1:1
    end

    @testset "setindex! range spans" begin
        layout = GridLayout(3, 4)
        layout[1:2, 2:3] = TextBox("Block")
        _, rows, cols = layout._entries[1]
        @test rows == 1:2
        @test cols == 2:3
    end

    @testset "setindex! bounds checking" begin
        layout = GridLayout(2, 2)
        @test_throws BoundsError (layout[3, 1] = TextBox(""))
        @test_throws BoundsError (layout[1, 0] = TextBox(""))
        @test_throws BoundsError (layout[1:3, 1] = TextBox(""))
    end

    @testset "setindex! conflict detection" begin
        layout = GridLayout(3, 3)
        layout[1, 1] = TextBox("A")
        # exact overlap
        @test_throws ArgumentError (layout[1, 1] = TextBox("B"))
        # partial overlap
        layout2 = GridLayout(3, 3)
        layout2[1:2, 1:2] = TextBox("Block")
        @test_throws ArgumentError (layout2[2, 2] = TextBox("Corner"))
        # non-overlapping should succeed
        layout3 = GridLayout(2, 2)
        layout3[1, 1] = TextBox("A")
        layout3[1, 2] = TextBox("B")
        layout3[2, :] = TextBox("C")
        @test length(layout3._entries) == 3
    end

    @testset "show" begin
        layout = GridLayout(2, 2)
        @test sprint(show, layout) == "GridLayout(2×2)"
        layout[1, 1] = TextBox("x")
        io = IOBuffer()
        Base.show(io, MIME"text/plain"(), layout)
        @test contains(String(take!(io)), "1 assigned cell")
    end

    @testset "push! assigns nested rids" begin
        slide = Slide()
        layout = GridLayout(1, 2)
        img1 = Picture(joinpath(PPTX.ASSETS_DIR, "julia_logo.emf"); size_x=10, size_y=10)
        img2 = Picture(joinpath(PPTX.ASSETS_DIR, "julia_dots.wmf"); size_x=10, size_y=10)
        layout[1, 1] = img1
        layout[1, 2] = img2

        push!(slide, layout)

        @test length(PPTX.shapes(slide)) == 1
        pushed_layout = only(PPTX.shapes(slide))
        @test pushed_layout isa GridLayout

        pushed_entries = pushed_layout._entries
        @test PPTX.rid(pushed_entries[1][1]) == 2
        @test PPTX.rid(pushed_entries[2][1]) == 3
        @test PPTX.rid(pushed_layout) == 3

        @test PPTX.new_rid(slide) == 4
    end

    @testset "nested relationships" begin
        target_slide = Slide()
        slide = Slide()
        layout = GridLayout(1, 2)
        img = Picture(joinpath(PPTX.ASSETS_DIR, "julia_logo.emf"); size_x=10, size_y=10)
        link_box = TextBox(content="jump", hlink=target_slide)
        layout[1, 1] = img
        layout[1, 2] = link_box
        push!(slide, layout)

        rel_map = PPTX.slide_relationship_map(slide)
        nested_img = only(filter(x -> x isa Picture, PPTX.shapes(layout)))
        @test haskey(rel_map, nested_img)
        @test haskey(rel_map, target_slide)

        rel_xml = PPTX.make_slide_relationships(slide)
        relationships = rel_xml["Relationships"]
        @test length(relationships) == 4
        @test relationships[3]["Relationship"][1]["Id"] == "rId2"
        @test relationships[4]["Relationship"][1]["Id"] == "rId3"
    end

    @testset "make_slide places grid shapes" begin
        slide = Slide()
        layout = GridLayout(2, 2; padding=10)
        layout[1, 1] = TextBox("Title")
        layout[2, :] = Picture(joinpath(PPTX.ASSETS_DIR, "julia_logo.emf"); size_x=10, size_y=10)
        push!(slide, layout)

        xml = PPTX.make_slide(
            slide;
            slide_size_x=PPTX.mm_to_emu(100),
            slide_size_y=PPTX.mm_to_emu(50),
        )

        sp_tree = xml["p:sld"][end]["p:cSld"][1]["p:spTree"]
        # spTree contains 2 group entries first, then our 2 layout shapes
        text_sp = sp_tree[3]["p:sp"]
        pic_sp = sp_tree[4]["p:pic"]

        text_xfrm = text_sp[2]["p:spPr"][1]["a:xfrm"]
        text_off = only(filter(x -> haskey(x, "a:off"), text_xfrm))["a:off"]
        text_ext = only(filter(x -> haskey(x, "a:ext"), text_xfrm))["a:ext"]
        @test text_off[1]["x"] == string(PPTX.mm_to_emu(10))
        @test text_off[2]["y"] == string(PPTX.mm_to_emu(10))
        @test text_ext[1]["cx"] == string(PPTX.mm_to_emu(35))
        @test text_ext[2]["cy"] == string(PPTX.mm_to_emu(10))

        pic_xfrm = pic_sp[3]["p:spPr"][1]["a:xfrm"]
        pic_off = pic_xfrm[1]["a:off"]
        pic_ext = pic_xfrm[2]["a:ext"]
        @test pic_off[1]["x"] == string(PPTX.mm_to_emu(10))
        @test pic_off[2]["y"] == string(PPTX.mm_to_emu(30))
        @test pic_ext[1]["cx"] == string(PPTX.mm_to_emu(80))
        @test pic_ext[2]["cy"] == string(PPTX.mm_to_emu(10))
    end
end
