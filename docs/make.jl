using PPTX
using Documenter
import Documenter.Remotes: GitHub

DocMeta.setdocmeta!(PPTX, :DocTestSetup, :(using PPTX, DataFrames, Colors); recursive=true)

makedocs(;
    modules=[PPTX],
    #doctest=false,
    authors="ASML",
    repo=GitHub("ASML-Labs", "PPTX.jl"),
    sitename="PPTX.jl",
    format=Documenter.HTML(;
        prettyurls=get(ENV, "CI", "false") == "true",
        edit_link="main",
        #assets=String[],
    ),
    pages=[
        "Home" => "index.md",
        "Table Styling" => "tablestyle.md",
        "Plots" => "plots.md",
        "API Reference" => "api.md",
    ],
)

deploydocs(
    repo = "github.com/ASML-Labs/PPTX.jl.git",
)
