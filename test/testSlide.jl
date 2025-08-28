@testset "Slide" begin
    slide = Slide()
    picture_path = joinpath(PPTX.ASSETS_DIR, "cauliflower.jpg")
    push!(slide, Picture(picture_path))
    @test slide.shapes[end].source == picture_path
end
