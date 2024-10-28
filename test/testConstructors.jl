using PPTX
import PPTX: slides, shapes, rid
using Test
using Colors

@testset "constructors" begin
    @testset "TextBox" begin
        box = TextBox("content")
        @test String(box.content) == "content"
        @test box.offset_x == Int(round(50*PPTX._EMUS_PER_MM))
        box = TextBox("content"; size_y = 80)
        @test String(box.content) == "content"
        @test box.size_y == Int(round(80*PPTX._EMUS_PER_MM))
        @test String(box.content) == "content"

        @test sprint(show, box) == "TextBox"

        box = TextBox("content"; linecolor = :black, linewidth = 3, color=:white)
        @test box.color == hex(colorant"white")
        @test box.linecolor == hex(colorant"black")
        @test box.linewidth == 38100 # EMUs

        # legacy dict style interface
        box = TextBox("content"; size_y = 80, style=Dict("italic" => true, "bold" => true, "fontsize" => 24.5))
        @test box.content.style.fontsize == 24.5
        @test box.content.style.italic == true
        @test box.content.style.bold == true

        t = TextStyle(fontsize = 24, italic = true, align=:center)
        @test t.align == "center"
        box2 = TextBox("content"; size_y = 80, style=t)
        @test box2.content.style == t

        c = colorant"red"
        t = TextStyle(color = :red)
        @test t.color == hex(c)
        PPTX.hex_color(t) == hex(c)

        args = (fontsize = 24, italic = true)
        box2 = TextBox("content"; size_y = 80, style=args)
        @test box2.content.style == TextStyle(; args...)

        io = IOBuffer()
        Base.show(io, MIME"text/plain"(), box)
        show_string = String(take!(io))
        @test contains(show_string, "content is \"$(String(box.content))\"")

        b = TextBox("bla", rotation=-90)
        @test b.rotation == 270.0

        b = TextBox("bla", rotation=360+45)
        @test b.rotation == 45.0
    end
    @testset "Picture" begin
        @test_throws ArgumentError pic = Picture("path")
        fnames = ["julia_logo.png", "julia_logo.svg", "julia_logo.emf", "julia_dots.wmf"]
        for fname in fnames
            logo_path = joinpath(PPTX.ASSETS_DIR, fname)
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
    
            contains(PPTX._show_string(pic2, false), "source is \"$(pic.source)\"")
        end
    end
    @testset "Picture - custom aspect ratio" begin
        logo_path = joinpath(PPTX.ASSETS_DIR,"julia_logo.svg")
        pic = Picture(logo_path; size_x=40, size_y=30)
        @test pic.size_x == 1440000
        @test pic.size_y == 1080000
    end
    @testset "empty" begin
        p = Presentation()
        ps = slides(p)
        @test length(ps) == 1
        @test ps[1].title == "My Presentation"

        s = Slide()
        @test isempty(shapes(s))
    end
    @testset "first slide" begin
        p = Presentation()
        @test rid(p.slides[1]) == 6
        picture_path = joinpath(PPTX.ASSETS_DIR, "cauliflower.jpg")
        p = Presentation([Slide([TextBox(),Picture(picture_path)])])
        @test rid(p.slides[1].shapes[1]) == 0
        @test rid(p.slides[1].shapes[2]) == 2
    end
    @testset "Slide" begin
        slide = Slide()
        picture_path = joinpath(PPTX.ASSETS_DIR, "cauliflower.jpg")
        push!(slide, Picture(picture_path))
    end
end

@testset "show AbstractShape" begin
    struct Something <: PPTX.AbstractShape end
    s = Something()
    @test PPTX.has_rid(s) == false
    PPTX.set_rid!(s, 2) # this does nothing
    @test PPTX.rid(s) == 0

    # default show
    @test sprint(show, s) == "Something"

    @testset "Presentation" begin
        p = Presentation()

        @test sprint(show, p) == "Presentation with 1 slide"

        io = IOBuffer()
        Base.show(io, MIME"text/plain"(), p)
        show_string = String(take!(io))
        @test contains(show_string, "title is \"$(p.title)\"")

        push!(p, Slide())
        @test sprint(show, p) == "Presentation with 2 slides"
    end
end
