using Test
using PPTX

@testset "Presentation Size" begin
    p = Presentation()
    template_reader = PPTX.ZipBufferReader(PPTX.read_template(PPTX.DEFAULT_TEMPLATE_DATA))
    PPTX.update_presentation_state!(p, template_reader)
    @test p._state.size.x == 12192000
    @test p._state.size.y == 6858000
end