# PPTX

[![Build Status](https://github.com/ASML-Labs/PPTX.jl/actions/workflows/CI.yml/badge.svg?branch=main)](https://github.com/ASML-Labs/PPTX.jl/actions/workflows/CI.yml?query=branch%3Amain)
[![Coverage](https://codecov.io/gh/ASML-Labs/PPTX.jl/branch/main/graph/badge.svg)](https://codecov.io/gh/ASML-Labs/PPTX.jl)
[![Code Style: Blue](https://img.shields.io/badge/code%20style-blue-4495d1.svg)](https://github.com/invenia/BlueStyle)
[![](https://img.shields.io/badge/docs-dev-blue.svg)](https://ASML-Labs/PPTX.jl/dev)

Package to generate PowerPoint© .pptx files.

Interface functions:
* `Presentation`: Presentation constructor, used for making a presentation.
* `Slide`: Slide constructor, used for making a slide.
* `TextBox`: TextBox constructor, used for adding text to slides.
* `Picture`: Picture constructor, used for adding pictures to slides.
* `write`: write a presentation to a file.

## Example usage

```julia
using PPTX, DataFrames

# Lets make a presentation
# note: this already adds a first slide with the title
pres = Presentation(; title="My First PowerPoint")

# What about a slide with some text
s1 = Slide(; title="My First Slide")
text = TextBox(; content="hello world!", offset_x=100, offset_y=100, size_x=150, size_y=20)
push!(s1, text)
text2 = TextBox(; content="here we are again", offset_x=100, offset_y=120, size_x=150, size_y=20)
push!(s1, text2)
push!(pres, s1)

# Now lets add a picture and some text
cauli_pic = Picture(joinpath(PPTX.EXAMPLE_DIR,"pictures/cauliflower.jpg"))
text = TextBox(content="Look its a vegetable!")
s3 = Slide()
push!(s3, cauli_pic)
push!(s3, text)

# move picture 100 mm down and 100 mm right
julia_logo = Picture(joinpath(PPTX.EXAMPLE_DIR,"pictures/julia_logo.png"), offset_x=100, offset_y=100)
push!(s3, julia_logo)
push!(pres, s3)

# and what about a table?
s4 = Slide(; title="A Table")
df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
my_table = Table(df; offset_x=60, offset_y=80, size_x=150, size_y=40)
push!(s4, my_table)
push!(pres, s4)

PPTX.write("example.pptx", pres, overwrite = true, open_ppt=true)
```