
# This struct exists for these reasons: 
# 1) To add all shapes in the layout as a flat vector in Slide (so id is same as index)
# 2) To ensure that the only contents of the layout can be AbstractShapes
# 3) To allow us to override setindex! with an AbstractShape without risk of ambiguity
# 4) to provide some better default values for the GridLayout (especially the bbox, see below)

# Main drawback is that we don't get all the utility functions defined for GridLayout for free since we must forward
# the calls. There are also a couple of functions which check x isa GridLayout internally which will not work
struct ShapeLayout
    slide::Slide
    layout::GridLayout
    function ShapeLayout(
        slide,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real; # millimeters
        kwargs...
        )
        offset_x_emu = offset_x * _EMUS_PER_MM
        offset_y_emu = offset_y * _EMUS_PER_MM
        size_x_emu = size_x * _EMUS_PER_MM
        size_y_emu = size_y * _EMUS_PER_MM

        # Maybe better for users to see this in mm, but all other shapes display in EMU
        suggestedbbox = BBox(
            offset_x_emu,               # Left
            offset_x_emu + size_x_emu,  # Right
            -offset_y_emu - size_y_emu, # Bottom. Negative since PPTX counts y from the top
            -offset_y_emu               # Top. Negative since PPTX counts y from the top
        )

        new(slide, GridLayout(;halign=:left, valign=:top, kwargs..., bbox = suggestedbbox))    
    end
end

# keyword argument constructor
function ShapeLayout(slide::Slide;
    offset_x::Real=20, # millimeters
    offset_y::Real=50, # millimeters
    size_x::Real=100, # millimeters
    size_y::Real=40, # millimeters
    kwargs... # Forwarded to GridLayout
    )
    # Would really like to get the actual layout of the slide here so we can initialize the bounding boxes
    ShapeLayout(slide, offset_x, offset_y, size_x, size_y; kwargs...)
end

ShapeLayout(sl::ShapeLayout; kwargs...) = ShapeLayout(sl.slide; kwargs...)

GridLayoutBase.layoutobservables(sl::ShapeLayout) = GridLayoutBase.layoutobservables(sl.layout)

# Make sure that the only things we can add are Shapes and ShapeLayouts
Base.setindex!(sl::ShapeLayout, content::ShapeLayout, args...) = setindex!(sl.layout, content, args...) 
function Base.setindex!(sl::ShapeLayout, content::AbstractShape, args...) 
    # Synopsis: Since shapes are stored as a flat vector outside of the layout we need to make sure it
    # stays in sync with the layout in case we overwrite. 
    # We do this by just storing the index inside the ShapeLayoutObservables. If there is something
    # at the position (given by args...) we just overwrite it instead of adding something new
    current = GridLayoutBase.contents(Base.getindex(sl.layout, args...)) 
    slo = if isempty(current)
        push!(sl.slide, content)
        ShapeLayoutObservables(content, lastindex(sl.slide.shapes))
    else
        shapeindex = only(current).index
        sl.slide.shapes[shapeindex] = content
        ShapeLayoutObservables(content, shapeindex)
    end
    setindex!(sl.layout, slo, args...)
end

Base.getindex(sl::ShapeLayout, args...) = Base.getindex(sl.layout, args...)

# This struct exists for these reasons:
# 1) the contents of a GridLayout must have a LayoutObservable
# 2) With the current design we need to remember which index we inserted the 
#    shape into slide.shapes since setindex! allows for overwriting

# Users should not need to interact with it
struct ShapeLayoutObservables{T<:AbstractShape}
    shape::T
    index::Int
    layoutobservables::LayoutObservables{GridLayout}
end

# TODO: Need to implement for each shape
BBox(s::AbstractShape) = BBox(Float32, 
                            s.offset_x[],               # Left 
                            s.offset_x[]+s.size_x[],    # Right  
                            -s.offset_y[] - s.size_y[], # Bottom. Negative since PPTX counts y from the top
                            -s.offset_y[]               # Top. Negative since PPTX counts y from the top
                            )

# Don't need to add LayoutObservables since GridLayout has it
ShapeLayoutObservables(sl::ShapeLayout, args...) = sl 
function ShapeLayoutObservables(shape::AbstractShape, shapeindex)

    # Just boilerplate to create a LayoutObservable
    # We only use the computedbboxobservable from all this for now.
    # Maybe we need to handle a few others for the full API of GridLayoutBase to work well...
    bbox = BBox(shape)
    layout_width = Observable{Any}(GridLayoutBase.width(bbox))
    layout_height = Observable{Any}(GridLayoutBase.height(bbox))
    layout_tellwidth = Observable(true)
    layout_tellheight = Observable(true)
    layout_halign = Observable{GridLayoutBase.HorizontalAlignment}(:left)
    layout_valign = Observable{GridLayoutBase.VerticalAlignment}(:top)
    layout_alignmode = Observable{Any}(GridLayoutBase.Inside())

    lobservables = LayoutObservables(
        layout_width,
        layout_height,
        layout_tellwidth,
        layout_tellheight,
        layout_halign,
        layout_valign,
        layout_alignmode,
    )

    sl = ShapeLayoutObservables(shape, shapeindex, lobservables)

    Observables.on(GridLayoutBase.computedbboxobservable(sl)) do newbbox
        updatebbox!(shape, newbbox)
    end
    # Would like to connect the sizes and offsets of shape to suggestedbboxobservable so users can change their size
    # and have the results take effect in the layout, but doing so create an infinite loop. Maybe it can be avoided through
    # with_updates_suspended
    sl
end

# TODO: Need to implement for each shape
function updatebbox!(s::AbstractShape, bbox)
    s.offset_x[] = round(Int, GridLayoutBase.left(bbox))
    s.offset_y[] = -round(Int, GridLayoutBase.top(bbox)) # Negative since PPTX counts y from the top
    s.size_x[] = round(Int, GridLayoutBase.width(bbox))
    s.size_y[] = round(Int, GridLayoutBase.height(bbox))
end
