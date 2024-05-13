using ZipArchives:
    ZipBufferReader, zip_readentry, zip_names, zip_append_archive, zip_newfile

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

@testset "writing" begin
    @testset "push same picture" begin
        picture_path = joinpath(PPTX.ASSETS_DIR, "cauliflower.jpg")
        p = Presentation([Slide([Picture(picture_path)]), Slide([Picture(picture_path)])])
        @test write_and_remove("test.pptx", p)
    end

    @testset "pushing same picture twice" begin
        pres = Presentation()
        s1 = Slide()
        julia_logo = Picture(joinpath(PPTX.ASSETS_DIR,"julia_logo.png"), top = 110, left = 110)
        push!(s1, julia_logo)
        push!(pres, s1)
        s2 = Slide()
        push!(s2, julia_logo)
        push!(pres, s2)

        @test write_and_remove("test.pptx", pres)
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

    # error testing
    wrong_path = abspath(joinpath(".", "wrong_path"))
    err_msg = "No file found at template path: $(repr(wrong_path))"
    @test_throws ErrorException(err_msg) PPTX.write(
        "anywhere.pptx",
        pres;
        overwrite=true,
        open_ppt=false,
        template_path=wrong_path,
    )

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
        output_zip = ZipBufferReader(read(output_pptx))
        dir_contents = zip_names(output_zip)
        theme_file = "ppt/theme/theme1.xml"
        @test theme_file ∈ dir_contents

        # file compare is failing on Linux, using string comparisons
        # the dark theme contains this node, which is named "<a:clrScheme name=\"Office\">" in the original theme
        str = zip_readentry(output_zip, theme_file, String)
        @test contains(str, "<a:clrScheme name=\"Office Theme\">")

        # make sure table styles are not empty
        @test PPTX.DEFAULT_TABLE_STYLE_DATA == zip_readentry(output_zip, "ppt/tableStyles.xml")
    end
end

@testset "custom template with media dir" begin
    # test for issue https://github.com/ASML-Labs/PPTX.jl/issues/20
    mktempdir() do tmpdir
        template_name = "no-slides.pptx"
        original_template_path = joinpath(PPTX.TEMPLATE_DIR, template_name)
        edited_template_path = joinpath(tmpdir, template_name)
        cp(original_template_path, edited_template_path)
        zip_append_archive(edited_template_path) do w
            # add an existing media directory
            zip_newfile(w, "ppt/media/foo.png")
            write(w, read(joinpath(PPTX.ASSETS_DIR,"julia_logo.png")))
        end

        pres = Presentation(;title="My Presentation")
        s1 = Slide()
        julia_logo = Picture(joinpath(PPTX.ASSETS_DIR,"julia_logo.png"), top = 110, left = 110)
        push!(s1, julia_logo)
        push!(pres, s1)

        # originally this threw IOError: mkdir("media"; mode=0o777): file already exists (EEXIST)
        pptx_path = joinpath(tmpdir, "example.pptx")
        write(pptx_path, pres; open_ppt=false, template_path=edited_template_path)

        @test isfile(pptx_path)
    end
end