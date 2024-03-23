
@testset "Test Hyperlinks" begin
    pres = Presentation()
    s2 = Slide()
    s3 = Slide()

    box = TextBox(content = "Hyperlinked text", hlink = s3)

    @testset "has hyperlink" begin
        @test PPTX.has_hyperlink(box)
        box_nolink = TextBox(content = "Hyperlinked text")
        @test !PPTX.has_hyperlink(box_nolink)
    end

    @testset "hyperlink xml" begin
        xml = PPTX.hlink_xml(box.hlink)
        @test xml["a:hlinkClick"]["rId"] == "rId$(PPTX.rid(s3))"
    end

    push!(s2, box)
    push!(pres, s2)
    push!(pres, s3)

    @testset "hyperlink xml after rid update" begin
        # After updating the rId of s2 the hyperlink xml should also be updated
        xml = PPTX.hlink_xml(box.hlink)
        @test xml["a:hlinkClick"]["rId"] == "rId$(PPTX.rid(s3))"
    end

    # I malificently swap slides which should not affect hyperlinking
    pres.slides[2] = s3
    pres.slides[3] = s2

    @testset "update slide nrs" begin
        PPTX.update_slide_nrs!(pres)
        @test pres.slides[2].slide_nr == 2
        @test pres.slides[3].slide_nr == 3
        xml_rels = PPTX.relationship_xml(box.hlink)
        # Chech that rId is still the same
        @test xml_rels["Relationship"][1]["Id"] == "rId$(PPTX.rid(s3))"
        @test xml_rels["Relationship"][3]["Target"] == "slide$(PPTX.slide_nr(s3)).xml"
    end

end