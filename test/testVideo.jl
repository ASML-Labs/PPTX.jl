@testset "Video" begin
    video_path = joinpath(PPTX.ASSETS_DIR, "sample_video.mp4")

    # Basic constructor
    v = Video(video_path)
    @test v.source == video_path
    @test v.offset_x == 0
    @test v.offset_y == 0
    @test v.size_x == Int(round(40 * PPTX._EMUS_PER_MM))
    @test v.size_y == Int(round(40 * PPTX._EMUS_PER_MM))
    @test v.rid == 0
    @test PPTX.has_rid(v) == true
    @test typeof(v._uuid) == String

    # Custom size and position
    v2 = Video(video_path; top=10, left=20, size_x=50, size_y=30)
    @test v2.offset_x == Int(round(20 * PPTX._EMUS_PER_MM))
    @test v2.offset_y == Int(round(10 * PPTX._EMUS_PER_MM))
    @test v2.size_x == Int(round(50 * PPTX._EMUS_PER_MM))
    @test v2.size_y == Int(round(30 * PPTX._EMUS_PER_MM))

    # RID setting
    v3 = PPTX.set_rid(v2, 5)
    @test v3.rid == 5
    @test rid(v3) == 5
    @test v3.source == v2.source
    @test v3.offset_x == v2.offset_x
    @test v3.offset_y == v2.offset_y
    @test v3.size_x == v2.size_x
    @test v3.size_y == v2.size_y

    # Show string
    show_str = PPTX._show_string(v3, false)
    @test occursin("source is", show_str)
    @test occursin("offset_x is", show_str)

    # Filename generation
    fname = PPTX.filename(v3)
    @test endswith(fname, ".mp4")

    # Relationship XML
    rel_xml = PPTX.relationship_xml(v3, 10)
    @test rel_xml["Relationship"][1]["Id"] == "rId10"
    @test occursin("media", rel_xml["Relationship"][3]["Target"])

    # Type schema
    @test PPTX.type_schema(v3) == "http://schemas.microsoft.com/office/2007/relationships/media"
    @test PPTX.type_schema(v3; it=1) == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"

    video_path = joinpath(PPTX.ASSETS_DIR, "sample_video.mp4")
    video = Video(video_path)
    @testset "make_xml" begin
        rel_map = Dict{Any, Int}(video => 2)
        xml = PPTX.make_xml(video, 1, rel_map)
        @test haskey(xml, "p:pic")
        @test length(xml["p:pic"]) == 3
    end
end

@testset "Thumbnail" begin
    video_path = joinpath(PPTX.ASSETS_DIR, "sample_video.mp4")
    v = Video(video_path)
    t = PPTX.create_thumbnail_image(v)

    @test t isa PermutedDimsArray
    
    t_name = PPTX.thumbnail_name(v)
    @test t_name == split(PPTX.filename(v), ".")[begin]*"_thumbnail.png"
end