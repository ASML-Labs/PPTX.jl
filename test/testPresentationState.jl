using Test
using PPTX
using Artifacts

@testset "Presentation Size" begin
    template_folder = abspath(joinpath(artifact"pptx_data", "templates","no-slides"))
    p = Presentation()
    ppt_dir = joinpath(template_folder, "ppt")
    PPTX.update_presentation_state!(p, ppt_dir)
    @test p._state.size.x == 12192000
    @test p._state.size.y == 6858000
end