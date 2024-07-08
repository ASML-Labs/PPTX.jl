
Base.@kwdef struct TextBody
    text::String
    style::AbstractDict = Dict("bold" => false, "italic" => false)
    body_properties::Union{Nothing, AbstractVector} = default_body_properties()
end

function TextBody(text::AbstractString; kwargs...)
    return TextBody(;text=convert(String, text), kwargs...)
end

function TextBody(text::AbstractString, style::AbstractDict)
    return TextBody(;text=convert(String, text), style=style)
end

function PlainTextBody(text::AbstractString)
    return TextBody(;text=convert(String, text), body_properties=nothing)
end

Base.convert(::Type{TextBody}, s::AbstractString) = TextBody(String(s))
Base.String(s::TextBody) = s.text

function default_body_properties()
    return [
        Dict("wrap" => "square"),
        Dict("rtlCol" => "0"),
        Dict("a:spAutoFit" => missing),
    ]
end

"""
```julia
TextBox(;
    content::String = "",
    offset_x::Real = 50,
    offset_y::Real = 50,
    size_x::Real = 40,
    size_y::Real = 30,
    style::Dict = Dict("bold" => false, "italic" => false),
)
```

A TextBox to be used on a Slide.
Offsets and sizes are in millimeters, but will be converted to EMU.

# Examples
```jldoctest
julia> using PPTX

julia> text = TextBox(content="Hello world!", size_x=30)
TextBox
 content is "Hello world!"
 offset_x is 1800000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1080000 EMUs
 size_y is 1080000 EMUs

```
"""
struct TextBox <: AbstractShape
    content::TextBody
    offset_x::Int # EMUs
    offset_y::Int # EMUs
    size_x::Int # EMUs
    size_y::Int # EMUs
    hlink::Union{Nothing, Any}
    function TextBox(
        content::AbstractString,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real, # millimeters
        style::Dict=Dict("bold" => false, "italic" => false),
        hlink::Union{Nothing, Any} = nothing
    )
        # input is in mm
        return new(
            TextBody(content, style),
            Int(round(offset_x * _EMUS_PER_MM)),
            Int(round(offset_y * _EMUS_PER_MM)),
            Int(round(size_x * _EMUS_PER_MM)),
            Int(round(size_y * _EMUS_PER_MM)),
            hlink
        )
    end
end

# keyword argument constructor
function TextBox(;
    content::AbstractString="",
    offset_x::Real=50, # millimeters
    offset_y::Real=50, # millimeters
    size_x::Real=40, # millimeters
    size_y::Real=30, # millimeters
    style::Dict=Dict("bold" => false, "italic" => false),
    hlink::Union{Nothing, Any} = nothing
)
    return TextBox(
        content,
        offset_x,
        offset_y,
        size_x,
        size_y,
        style,
        hlink
    )
end

TextBox(content::String; kwargs...) = TextBox(;content=content, kwargs...)

function _show_string(p::TextBox, compact::Bool)
    show_string = "TextBox"
    if !compact
        show_string *= "\n content is \"$(String(p.content))\""
        show_string *= "\n offset_x is $(p.offset_x) EMUs"
        show_string *= "\n offset_y is $(p.offset_y) EMUs"
        show_string *= "\n size_x is $(p.size_x) EMUs"
        show_string *= "\n size_y is $(p.size_y) EMUs"
    end
    return show_string
end

function text_style_xml(t::TextBody)
    style = [Dict("lang" => "en-US")]
    if t.style["bold"]
        push!(stile, Dict("b" => "1"))
    end

    if t.style["italic"]
        push!(style, Dict("i" => "1"))
    end

    push!(style, Dict("dirty" => "0"))
    return style
end

function make_xml(t::TextBox, id::Int=1)
    cNvPr = Dict("p:cNvPr" => Dict[Dict("id" => "$id"), Dict("name" => "TextBox")])
    if has_hyperlink(t)
        push!(cNvPr["p:cNvPr"], hlink_xml(t.hlink))
    end

    cNvSpPr = Dict("p:cNvSpPr" => Dict("txBox" => "1"))
    nvPr = Dict("p:nvPr" => missing)

    nvSpPr = Dict("p:nvSpPr" => [cNvPr, cNvSpPr, nvPr])

    offset = Dict("a:off" => [Dict("x" => "$(t.offset_x)"), Dict("y" => "$(t.offset_y)")])
    extend = Dict("a:ext" => [Dict("cx" => "$(t.size_x)"), Dict("cy" => "$(t.size_y)")])

    spPr = Dict(
        "p:spPr" => [
            Dict("a:xfrm" => [offset, extend]),
            Dict("a:prstGeomt" => [Dict("prst" => "rect"), Dict("a:avLst" => missing)]),
            Dict("a:noFill" => missing),
        ],
    )

    txBody = make_textbody_xml(t)

    return Dict("p:sp" => [nvSpPr, spPr, txBody])
end

function make_textbody_xml(t::TextBody, txBodyNameSpace="p")
    txBody = Dict(
        "$txBodyNameSpace:txBody" => [
            Dict(
                "a:bodyPr" => t.body_properties,
            ),
            Dict("a:lstStyle" => missing),
            Dict(
                "a:p" => Dict(
                    "a:r" => [Dict("a:rPr" => text_style_xml(t)), Dict("a:t" => t)],
                ),
            ),
        ],
    )
    return txBody
end

make_textbody_xml(t::TextBox) = make_textbody_xml(t.content)