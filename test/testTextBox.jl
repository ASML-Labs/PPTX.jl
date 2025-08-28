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

        b = TextBox("bla", margins = (left=0.1, right=0.1))
        @test b.content.margins.left == 36000
        @test b.content.margins.right == 36000
        @test b.content.margins.top === nothing
        @test b.content.margins.bottom === nothing
    end