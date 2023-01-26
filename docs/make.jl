using PPTX
using Documenter

DocMeta.setdocmeta!(PPTX, :DocTestSetup, :(using PPTX); recursive=true)

makedocs(;
    modules=[PPTX],
    authors="ASML",
    repo="https://github.com/ASML-Labs/PPTX.jl/blob/{commit}{path}#{line}",
    sitename="PPTX.jl",
    format=Documenter.HTML(;
        prettyurls=get(ENV, "CI", "false") == "true",
        edit_link="main",
        assets=String[],
    ),
    pages=[
        "Home" => "index.md",
        "API Reference" => "api.md",
    ],
)

deploydocs(
    repo = "github.com/ASML-Labs/PPTX.jl.git",
)
