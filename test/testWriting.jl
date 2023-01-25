
function bincompare(path::String, ref::String)
    bin1 = read(path)
    bin2 = read(ref)
    same = true
    i = 1
    for (b1,b2) in zip(bin1,bin2)
        same = same & (b1 == b2)
        if !same
            println("$b1 is not $b2 at byte $i")
            break
        end
        i+=1
    end
    return same
end

function write_and_remove(fname::String, p::Presentation)
    PPTX.write(fname, p, overwrite = true, open_ppt = false)
    rm(fname)
    return true
end

@testset "zipping/unzipping" begin
    origin = @__DIR__

    no_slides_template = abspath(joinpath(PPTX.TEMPLATE_DIR,"no-slides.pptx"))
    mktempdir() do tmpdir
        cd(tmpdir)
        cp(no_slides_template, abspath(joinpath(".","no-slides-cp.pptx")))
        PPTX.unzip("no-slides-cp.pptx")
        dir_contents = readdir("no-slides-cp")
        @test "[Content_Types].xml" ∈ dir_contents
        @test "_rels" ∈ dir_contents
        @test "docProps" ∈ dir_contents
        @test "ppt" ∈ dir_contents
        cd(origin)
    end

    target = abspath(joinpath(PPTX.TESTDATA_DIR,"rezipped.pptx"))
    if isfile(target)
        rm(target)
    end
    mktempdir() do tmpdir
        cd(tmpdir)
        cp(no_slides_template, abspath(joinpath(".","no-slides.pptx")))
        PPTX.unzip("no-slides.pptx")
        PPTX.zip("no-slides","rezipped.pptx")
        # cp("rezipped.pptx", target)
        cd(origin)
    end
    # @test bincompare(no_slides_template, target)

    @testset "special cases" begin
        @testset "push same picture" begin
            picture_path = joinpath(PPTX.EXAMPLE_DIR, "pictures", "cauliflower.jpg")
            p = Presentation([Slide([Picture(picture_path)]), Slide([Picture(picture_path)])])
            @test write_and_remove("test.pptx", p)
        end

        @testset "pushing same picture twice" begin
            pres = Presentation()
            s1 = Slide()
            julia_logo = Picture(joinpath(PPTX.EXAMPLE_DIR,"pictures/julia_logo.png"), top = 110, left = 110)
            push!(s1, julia_logo)
            push!(pres, s1)
            s2 = Slide()
            push!(s2, julia_logo)
            push!(pres, s2)

            @test write_and_remove("test.pptx", pres)

        end
    end

end

# TODO: also support .potx next to empty .pptx
# TODO: what if the .pptx template already has slides?
@testset "custom template" begin
    dark_template_name = "no-slides-dark.pptx"
    dark_template_path = joinpath(PPTX.TEMPLATE_DIR, dark_template_name)
    pres = Presentation(;title="My Presentation")
    s = Slide()
    push!(pres, s)
    mktempdir() do tmpdir
        filename = "testfile-dark"
        output_pptx = abspath(joinpath(tmpdir, "$filename.pptx"))
        PPTX.write(
            output_pptx,
            pres;
            overwrite=true,
            open_ppt=false,
            template_path=dark_template_path,
        )
        PPTX.unzip(output_pptx)
        output_unzipped_pptx = abspath(joinpath(tmpdir, filename))
        dir_contents = readdir(output_unzipped_pptx)
        theme_file = joinpath(output_unzipped_pptx, "ppt", "theme", "theme1.xml")
        @test isfile(theme_file)

        # file compare is failing on Linux, using string comparisons
        # the dark theme contains this node, which is named "<a:clrScheme name=\"Office\">" in the original theme
        str = read(theme_file, String)
        @test contains(str, "<a:clrScheme name=\"Office Theme\">")
    end
end
