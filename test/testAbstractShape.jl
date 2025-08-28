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
