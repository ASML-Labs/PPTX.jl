using PPTX
using Test
using Artifacts

dark_template_name = "no-slides-dark.pptx"
dark_template_path = joinpath(artifact"pptx_data", "templates", dark_template_name)
isfile(dark_template_path)
pres = Presentation(;title="My Presentation")
s = Slide()
push!(pres, s)