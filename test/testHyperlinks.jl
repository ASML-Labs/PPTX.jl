
@testset "Test Hyperlinks" begin
    pres = Presentation()
    s2 = Slide()
    s3 = Slide()
    s4 = Slide()

    box = TextBox(content = "Hyperlinked text", hlink = s3)
    push!(s2, box)

    @testset "has hyperlink" begin
        @test PPTX.has_hyperlink(box)
        box_nolink = TextBox(content = "Hyperlinked text")
        @test !PPTX.has_hyperlink(box_nolink)
    end

    @testset "hyperlink xml" begin
        relationship_map = PPTX.slide_relationship_map(s2)
        @test relationship_map[box.hlink] == 2
        xml = PPTX.hlink_xml(box.hlink, relationship_map)
        @test xml["a:hlinkClick"]["r:id"] == "rId2"
    end

    # push the same link once more
    box2 = TextBox(content = "Hyperlinked text", hlink = s3)
    push!(s2, box2)

    @testset "duplicate hyperlink relations" begin
        relationship_map = PPTX.slide_relationship_map(s2)
        @test length(relationship_map) == 1
        @test relationship_map[box2.hlink] == 2
    end

    box3 = TextBox(content = "Hyperlinked text", hlink = s4)
    push!(s2, box3)

    @testset "multiple hyperlink relations" begin
        relationship_map = PPTX.slide_relationship_map(s2)
        @test length(relationship_map) == 2
        @test relationship_map[box.hlink] == 2
        @test relationship_map[box2.hlink] == 2
        @test relationship_map[box3.hlink] == 3
    end

end