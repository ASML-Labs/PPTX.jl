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

@testset "Picture - Thumbnail" begin
    logo_path = joinpath(PPTX.ASSETS_DIR,"julia_logo.svg")
    pic_thumbnail = PPTX.picture_thumbnail(logo_path)
    @test pic_thumbnail._uuid == ""
end