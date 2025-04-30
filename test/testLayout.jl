using PPTX
using Test

@testset "PPTX Layout names listing" begin
    layoutmap = PPTX.list_layoutnames()
    @test length(keys(layoutmap)) == 11    
    @test layoutmap["Vertical Title and Text"] == 11
end

@testset "PPTX Layout names writing" begin
    pres = Presentation()
    s2 = Slide(; layout="Comparison")
    push!(pres, s2)
    s3 = Slide(; layout="Section Header")
    push!(pres, s3)

    @test pres.slides[2].layout == "Comparison"
    @test pres.slides[3].layout == "Section Header"

    mktempdir() do tmpdir
        filename = "testfile-layout"
        output_pptx = abspath(joinpath(tmpdir, "$filename.pptx"))
        PPTX.write(output_pptx, pres)
    end

    @test pres.slides[2].layout == 5
    @test pres.slides[3].layout == 3
end

@testset "PPTX Layout not defined" begin

    @testset "Number not defined" begin
        pres = Presentation()
        s2 = Slide(; layout=12)
        push!(pres, s2)

        mktempdir() do tmpdir
            filename = "testfile-layout_number_not_defined"
            output_pptx = abspath(joinpath(tmpdir, "$filename.pptx"))
            err_msg = "Slide layout number 12 not defined in the template"
            @test_throws ErrorException(err_msg)  PPTX.write(output_pptx, pres)
        end
    end

    @testset "Name not defined" begin
        pres = Presentation()
        s2 = Slide(; layout="Three Content")
        push!(pres, s2)

        mktempdir() do tmpdir
            filename = "testfile-layout_name_not_defined"
            output_pptx = abspath(joinpath(tmpdir, "$filename.pptx"))
            err_msg = "Slide layout name Three Content not defined in the template"
            @test_throws ErrorException(err_msg)  PPTX.write(output_pptx, pres)
        end
    end

end
