using PPTX
using Test

@testset "constructors" begin
    @testset "TextBox" begin
        box = TextBox("content")
        @test String(box.content) == "content"
        @test box.offset_x == Int(round(50*PPTX._EMUS_PER_MM))
        box = TextBox("content"; size_y = 80)
        @test String(box.content) == "content"
        @test box.size_y == Int(round(80*PPTX._EMUS_PER_MM))
        box = TextBox(content="content", size_y = 80)
        @test String(box.content) == "content"
    end
    @testset "Picture" begin
        @test_throws ArgumentError pic = Picture("path")
        logo_path = joinpath(PPTX.EXAMPLE_DIR,"pictures/julia_logo.png")
        pic = Picture(logo_path)
        @test pic.offset_x == 0
        width = 150
        pic = Picture(logo_path, top = 100, left = 120, size=width)
        @test pic.offset_x == Int(round(120*PPTX._EMUS_PER_MM))
        @test pic.offset_y == Int(round(100*PPTX._EMUS_PER_MM))
        @test pic.size_x == Int(round(width*PPTX._EMUS_PER_MM))
        ratio = PPTX.image_aspect_ratio(logo_path)
        height = Int(round(width / ratio*PPTX._EMUS_PER_MM))
        @test pic.size_y == height

        # make setting the rid doesn't change the other values
        pic2 = PPTX.set_rid(pic, 5)
        @test pic2.rid == 5
        @test pic2.source == pic.source
        @test pic2.offset_x == pic.offset_x
        @test pic2.offset_y == pic.offset_y
        @test pic2.size_x == pic.size_x
        @test pic2.size_y == pic.size_y
    end
    @testset "empty" begin
        p = Presentation()
        ps = slides(p)
        @test length(ps) == 1
        @test ps[1].title == "unknown"

        s = Slide()
        @test isempty(shapes(s))
    end
    @testset "first slide" begin
        p = Presentation([Slide()])
        @test rid(p.slides[1]) == 6
        picture_path = joinpath(PPTX.EXAMPLE_DIR, "pictures", "cauliflower.jpg")
        p = Presentation([Slide([TextBox(),Picture(picture_path)])])
        @test rid(p.slides[1].shapes[1]) == 0
        @test rid(p.slides[1].shapes[2]) == 1
    end
    @testset "Slide" begin
        slide = Slide()
        picture_path = joinpath(PPTX.EXAMPLE_DIR, "pictures", "cauliflower.jpg")
        push!(slide, Picture(picture_path))
    end
end