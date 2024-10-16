```@meta
CurrentModule = PPTX
```

# PPTX

Documentation for [PPTX](https://github.com/ASML-Labs/PPTX.jl).

Interface functions are:
```@index
Pages   = ["api.md"]
```

## Example usage

You can build a presentation inside Julia:

```jldoctest
using PPTX, DataFrames

# Lets make a presentation
# note: this already adds a first slide with the title
pres = Presentation(; title="My First PowerPoint")

# What about a slide with some text
s2 = Slide(; title="My First Slide")
text = TextBox(; content="hello world!", offset_x=100, offset_y=100, size_x=150, size_y=20)
push!(s2, text)
text2 = TextBox(;
    content="here we are again",
    offset_x=100, offset_y=120, size_x=150, size_y=20,
    style=(italic=true, fontsize=16)
)
push!(s2, text2)
push!(pres, s2)

# Now lets add a picture and some text
cauli_pic = Picture(joinpath(PPTX.ASSETS_DIR,"cauliflower.jpg"))
text = TextBox(content="Look its a vegetable!")
s3 = Slide()
push!(s3, cauli_pic)
push!(s3, text)

# move picture 100 mm down and 100 mm right
julia_logo = Picture(joinpath(PPTX.ASSETS_DIR,"julia_logo.png"), offset_x=100, offset_y=100)
push!(s3, julia_logo)
push!(pres, s3)

# and what about a table?
s4 = Slide(; title="A Table")
df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
my_table = Table(df; offset_x=60, offset_y=80, size_x=150, size_y=40)
push!(s4, my_table)
push!(pres, s4)

# and what about a nice link in slide 2 to the table-slide
text = TextBox(;
    content="Click here to see a nice table",
    offset_x=100, offset_y=140, size_x=150, size_y=20,
    hlink = s4 # link to slide 4
)
push!(s2, text)

pres

# output

Presentation with 4 slides
 title is "My First PowerPoint"
 author is "unknown"

```

Finally you can write the PPTX file with `PPTX.write`:

```julia
PPTX.write("example.pptx", pres, overwrite = true, open_ppt=true)
```
