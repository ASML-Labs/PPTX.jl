# Table styling

A few examples on how to style a table.

## Default table

A table/dataframe with regular values will use the default .pptx table style:

```julia
using PPTX, DataFrames

p = Presentation()
s = Slide(title="table example")
push!(p, s)

df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
t = Table(df; offset_x=50, offset_y=50, size_x=80, size_y=60)

push!(s, t)

write("example.pptx", p; overwrite=true)
```

```@raw html
<img src="../images/table.png" width="400px"/>
```

## General table styles

You can disable the `header` (automatically generated from the dataframe column names) and the `bandrow` (the 'banded' or alternating coloring of different rows).

```julia
using PPTX, DataFrames

p = Presentation()
s = Slide(title="table example")
push!(p, s)

df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
t = Table(df;
    header=false, bandrow=false,
    offset_x=50, offset_y=50, size_x=80, size_y=40
)

push!(s, t)

write("example.pptx", p; overwrite=true)
```

```@raw html
<img src="../images/table_no_header.png" width="400px"/>
```

## Table element styling

You can overrule the table style per element.

For example, let's color one box and add lines around it. We'll also center the alignment and anchor (vertical alignment) of the text.

And let's also change the `dash` value. Possible values are "solid", "dot", "dash", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDashDotDot". Default is "solid".

You may have to fiddle a little with the lines, since neighboring cells have priority.

```julia
using PPTX, Colors, DataFrames

p = Presentation()
s = Slide(title="table example")
push!(p, s)

# a red cell with lines on all edges
# top and left are actually overruled by neighboring cells
t4 = TableCell(
    4;
    color=colorant"darkred",
    textstyle=(color=colorant"white",align="center"),
    anchor="center",
    lines=(
        bottom=(width=3,color=colorant"black"),
        right=(width=3,color=colorant"black",dash="sysDash"),
    )
)

# define lines of surrounding elements

t2 = TableCell(
    2;
    lines=(
        right=(width=3,color=colorant"black",dash="sysDash"),
    )
)

t3 = TableCell(
    3;
    lines=(
        bottom=(width=3,color=colorant"black"),
    )
)

df = DataFrame(a = [1,t2], b = [t3,t4], c = [5,6])
t = Table(df; offset_x=50, offset_y=50, size_x=80, size_y=60)

push!(s, t)

write("example.pptx", p; overwrite=true)
```

```@raw html
<img src="../images/styled_table.png" width="400px"/>
```
