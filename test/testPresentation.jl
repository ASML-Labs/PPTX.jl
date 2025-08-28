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