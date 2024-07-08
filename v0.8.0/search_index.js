var documenterSearchIndex = {"docs":
[{"location":"api/#API-reference","page":"API Reference","title":"API reference","text":"","category":"section"},{"location":"api/","page":"API Reference","title":"API Reference","text":"Modules = [PPTX]","category":"page"},{"location":"api/","page":"API Reference","title":"API Reference","text":"Modules = [PPTX]","category":"page"},{"location":"api/#PPTX.Picture","page":"API Reference","title":"PPTX.Picture","text":"Picture(source::String; top::Int=0, left::Int=0, size::Int = 40)\n\nsource::String path of image file\ntop::Int mm from the top\nleft::Int mm from the left\n\nInternally the sizes are converted EMUs.\n\nExamples\n\njulia> using PPTX\n\njulia> img = Picture(joinpath(PPTX.ASSETS_DIR, \"cauliflower.jpg\"))\nPicture\n source is \"./cauliflower.jpg\"\n offset_x is 0 EMUs\n offset_y is 0 EMUs\n size_x is 1440000 EMUs\n size_y is 1475072 EMUs\n\n\nOptionally, you can set the size_x and size_y manually for filetypes not supported by FileIO, such as SVG.\n\njulia> using PPTX\n\njulia> img = Picture(joinpath(PPTX.ASSETS_DIR, \"julia_logo.svg\"); size_x=40, size_y=30)\nPicture\n source is \"./julia_logo.svg\"\n offset_x is 0 EMUs\n offset_y is 0 EMUs\n size_x is 1440000 EMUs\n size_y is 1080000 EMUs\n\n\n\n\n\n\n","category":"type"},{"location":"api/#PPTX.Presentation","page":"API Reference","title":"PPTX.Presentation","text":"Presentation(\n    slides::Vector{Slide}=Slide[];\n    title::String=\"unknown\",\n    author::String=\"unknown\",\n)\n\nType to contain the final presentation you want to write to .pptx.\n\nIf isempty(slides) then we add a first slide with the Title slide layout.\n\nExamples\n\njulia> using PPTX\n\njulia> pres = Presentation(; title = \"My Presentation\")\nPresentation with 1 slide\n title is \"My Presentation\"\n author is \"unknown\"\n\n\n\n\n\n\n","category":"type"},{"location":"api/#PPTX.Slide","page":"API Reference","title":"PPTX.Slide","text":"Slide(\n    shapes::Vector{AbstractShape}=AbstractShape[];\n    title::String=\"\",\n    layout::Int=1,\n)\n\nshapes::Vector{AbstractShape} shapes to add to the PowerPoint, can also be pushed afterwards\ntitle::String title text placed inside the title textbox found in the slide layout\nlayout::Int which slide layout to use. Typically 1 is the title slide and 2 is the text slide.\n\nMake a Slide for a powerpoint Presentation.\n\nYou can push! any AbstractShape types into this slide, such as a TextBox or Picture.\n\nExamples\n\njulia> using PPTX\n\njulia> slide = Slide(; title=\"Hello Title\", layout=2)\nSlide(\"Hello Title\", PPTX.AbstractShape[], 0, 2)\n\njulia> text = TextBox(\"Hello world!\")\nTextBox\n content is \"Hello world!\"\n offset_x is 1800000 EMUs\n offset_y is 1800000 EMUs\n size_x is 1440000 EMUs\n size_y is 1080000 EMUs\n\njulia> push!(slide, text);\n\njulia> slide\nSlide(\"Hello Title\", PPTX.AbstractShape[TextBox], 0, 2)\n\n\n\n\n\n\n","category":"type"},{"location":"api/#PPTX.Table","page":"API Reference","title":"PPTX.Table","text":"Table(;\n    content,\n    offset_x::Real = 50,\n    offset_y::Real = 50,\n    size_x::Real = 150,\n    size_y::Real = 100,\n)\n\nA Table to be used on a Slide.\n\nThe content can be anything that adheres to a Tables.jl interface.\n\nOffsets and sizes are in millimeters, but will be converted to EMU.\n\nExamples\n\njulia> using PPTX, DataFrames\n\njulia> df = DataFrame(a = [1,2], b = [3,4], c = [5,6])\n2×3 DataFrame\n Row │ a      b      c     \n     │ Int64  Int64  Int64 \n─────┼─────────────────────\n   1 │     1      3      5\n   2 │     2      4      6\n\njulia> t = Table(content=df, size_x=30)\nTable\n content isa DataFrames.DataFrame\n offset_x is 1800000 EMUs\n offset_y is 1800000 EMUs\n size_x is 1080000 EMUs\n size_y is 3600000 EMUs\n\n\n\n\n\n\n","category":"type"},{"location":"api/#PPTX.TextBox","page":"API Reference","title":"PPTX.TextBox","text":"TextBox(;\n    content::String = \"\",\n    offset_x::Real = 50,\n    offset_y::Real = 50,\n    size_x::Real = 40,\n    size_y::Real = 30,\n    style::Dict = Dict(\"bold\" => false, \"italic\" => false),\n)\n\nA TextBox to be used on a Slide. Offsets and sizes are in millimeters, but will be converted to EMU.\n\nExamples\n\njulia> using PPTX\n\njulia> text = TextBox(content=\"Hello world!\", size_x=30)\nTextBox\n content is \"Hello world!\"\n offset_x is 1800000 EMUs\n offset_y is 1800000 EMUs\n size_x is 1080000 EMUs\n size_y is 1080000 EMUs\n\n\n\n\n\n\n","category":"type"},{"location":"api/#Base.write-Tuple{String, Presentation}","page":"API Reference","title":"Base.write","text":"Base.write(\n    filepath::String,\n    p::Presentation;\n    overwrite::Bool=false,\n    open_ppt::Bool=true,\n    template_path::String=\"no-slides.pptx\",\n)\n\nfilepath::String Desired presentation filepath.\npres::Presentation Presentation object to write.\noverwrite = false argument for overwriting existing file.\nopen_ppt = true open powerpoint after it is written.\ntemplate_path::String path to an (empty) pptx that serves as template.\n\nExamples\n\njulia> using PPTX\n\njulia> slide = Slide()\n\njulia> text = TextBox(\"Hello world!\")\n\njulia> push!(slide, text)\n\njulia> pres = Presentation()\n\njulia> push!(pres, slide)\n\njulia> write(\"hello_world.pptx\", pres)\n\n\n\n\n\n","category":"method"},{"location":"","page":"Home","title":"Home","text":"CurrentModule = PPTX","category":"page"},{"location":"#PPTX","page":"Home","title":"PPTX","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"Documentation for PPTX.","category":"page"},{"location":"","page":"Home","title":"Home","text":"Interface functions are:","category":"page"},{"location":"","page":"Home","title":"Home","text":"Pages   = [\"api.md\"]","category":"page"},{"location":"#Example-usage","page":"Home","title":"Example usage","text":"","category":"section"},{"location":"","page":"Home","title":"Home","text":"You can build a presentation inside Julia:","category":"page"},{"location":"","page":"Home","title":"Home","text":"using PPTX, DataFrames\n\n# Lets make a presentation\n# note: this already adds a first slide with the title\npres = Presentation(; title=\"My First PowerPoint\")\n\n# What about a slide with some text\ns2 = Slide(; title=\"My First Slide\")\ntext = TextBox(; content=\"hello world!\", offset_x=100, offset_y=100, size_x=150, size_y=20)\npush!(s2, text)\ntext2 = TextBox(; content=\"here we are again\", offset_x=100, offset_y=120, size_x=150, size_y=20)\npush!(s2, text2)\npush!(pres, s2)\n\n# Now lets add a picture and some text\ncauli_pic = Picture(joinpath(PPTX.ASSETS_DIR,\"cauliflower.jpg\"))\ntext = TextBox(content=\"Look its a vegetable!\")\ns3 = Slide()\npush!(s3, cauli_pic)\npush!(s3, text)\n\n# move picture 100 mm down and 100 mm right\njulia_logo = Picture(joinpath(PPTX.ASSETS_DIR,\"julia_logo.png\"), offset_x=100, offset_y=100)\npush!(s3, julia_logo)\npush!(pres, s3)\n\n# and what about a table?\ns4 = Slide(; title=\"A Table\")\ndf = DataFrame(a = [1,2], b = [3,4], c = [5,6])\nmy_table = Table(df; offset_x=60, offset_y=80, size_x=150, size_y=40)\npush!(s4, my_table)\npush!(pres, s4)\n\n# and what about a nice link in slide 2 to the table-slide\ntext = TextBox(; content=\"Click here to see a nice table\", offset_x=100, offset_y=140, size_x=150, size_y=20, hlink = s4)\npush!(s2, text)\n\npres\n\n# output\n\nPresentation with 4 slides\n title is \"My First PowerPoint\"\n author is \"unknown\"\n","category":"page"},{"location":"","page":"Home","title":"Home","text":"Finally you can write the PPTX file with PPTX.write:","category":"page"},{"location":"","page":"Home","title":"Home","text":"PPTX.write(\"example.pptx\", pres, overwrite = true, open_ppt=true)","category":"page"}]
}
