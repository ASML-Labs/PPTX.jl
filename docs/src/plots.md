# Plots

Quick examples for how make a plot, save it as a figure and add it to a .pptx presentation.

## Plots.jl

```julia
using PPTX, Plots

p = Presentation(title="Plots", author="PPTX.jl")
s = Slide(title="Plots.jl example")
push!(p, s)

# create a save a
x = range(0, 10, length=100)
y = sin.(x)
f = plot(x, y, size=(600,400))
Plots.savefig(f, "my_plot.png")

my_picture = Picture("my_plot.png", offset = (50,50), size = 160)
push!(s, my_picture)

t = TextBox("Code used:", offset = (210,70), textstyle = (bold=true,))
push!(s, t)

text = """
x = range(0, 10, length=100)
y = sin.(x)
f = plot(x, y)
"""
t = TextBox(text, offset = (215,80), size = (110,27), color = :lightgrey, textstyle = (typeface = "Courier New",))
push!(s, t)

write("example.pptx", p; overwrite=true)
```

## Makie.jl

Choose your own Makie backend (GLMakie, CairoMakie, etc).

```julia
using PPTX, GLMakie, Makie

p = Presentation(title="Plots", author="PPTX.jl")
s = Slide(title="Makie.jl example")
push!(p, s)

# create and save a figure
f = Figure(size=(600,400))
ax = Axis(f[1, 1])
x = range(0, 10, length=100)
y = sin.(x)
lines!(ax, x, y)
save("my_makie_plot.png", f)

my_picture = Picture("my_makie_plot.png", offset = (50,50), size = 160)
push!(s, my_picture)

t = TextBox("Code used:", offset = (210,70), textstyle = (bold=true,))
push!(s, t)

text = """
f = Figure(size=(600,400))
ax = Axis(f[1, 1])
x = range(0, 10, length=100)
y = sin.(x)
lines!(ax, x, y)
"""
t = TextBox(text, offset = (215,80), size = (110,45), color = :lightgrey, textstyle = (typeface = "Courier New",))
push!(s, t)

write("example.pptx", p; overwrite=true)

```