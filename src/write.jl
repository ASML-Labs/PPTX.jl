import DefaultApplication

function write_presentation!(w::ZipWriter, p::Presentation)
    xml = make_presentation(p)
    doc = xml_document(xml)
    zip_newfile(w, "ppt/presentation.xml"; compress=true)
    print(w, doc)
    zip_commitfile(w)
end

function write_relationships!(w::ZipWriter, p::Presentation)
    xml = make_relationships(p)
    doc = xml_document(xml)
    zip_newfile(w, "ppt/_rels/presentation.xml.rels"; compress=true)
    print(w, doc)
    zip_commitfile(w)
end

function write_slides!(w::ZipWriter, p::Presentation, template::ZipBufferReader)
    if zip_isdir(template, "ppt/slides")
        error("input template pptx already contains slides, please use an empty template")
    end
    for (idx, slide) in enumerate(slides(p))
        xml = make_slide(slide)
        doc::EzXML.Document = xml_document(xml)
        add_title_shape!(doc, slide, template)
        zip_newfile(w, "ppt/slides/slide$idx.xml"; compress=true)
        print(w, doc)
        zip_commitfile(w)
        xml = make_slide_relationships(slide)
        doc = xml_document(xml)
        zip_newfile(w, "ppt/slides/_rels/slide$idx.xml.rels"; compress=true)
        print(w, doc)
        zip_commitfile(w)
    end
end

function add_title_shape!(doc::EzXML.Document, slide::Slide, template::ZipBufferReader)
    # xpath to find something with an unregistered namespace
    spTree = findfirst("//*[local-name()='p:spTree']", root(doc))
    layout_path = "ppt/slideLayouts/slideLayout$(slide.layout).xml"
    layout_doc = EzXML.parsexml(zip_readentry(template, layout_path))
    title_shape_node = PPTX.get_title_shape_node(layout_doc)
    if !isnothing(title_shape_node)
        PPTX.update_xml_title!(title_shape_node, slide.title)
        new_id = maximum(get_shape_ids(doc))+1
        update_shape_id!(title_shape_node, new_id)
        unlink!(title_shape_node)
        link!(spTree, title_shape_node)
    end
    nothing
end

function write_shapes!(w::ZipWriter, pres::Presentation)
    for slide in slides(pres)
        for shape in shapes(slide)
            if shape isa Picture
                copy_picture(w::ZipWriter, shape)
            end
        end
    end
end

function update_table_style!(w::ZipWriter, template::ZipBufferReader)
    # minimally we want at least one table style
    table_style_path = "ppt/tableStyles.xml"
    table_style_doc = EzXML.parsexml(zip_readentry(template, table_style_path))
    if has_empty_table_list(table_style_doc)
        zip_newfile(w, table_style_path; compress=true)
        write(w, DEFAULT_TABLE_STYLE_DATA)
        zip_commitfile(w)
    end
    nothing
end

function add_contenttypes!(w::ZipWriter, template::ZipBufferReader)
    path = "[Content_Types].xml"
    doc = EzXML.parsexml(zip_readentry(template, path))
    r = root(doc)
    extension_contenttypes = (
        ("emf", "image/x-emf"),
        ("gif", "image/gif"),
        ("jpeg", "image/jpeg"),
        ("jpg", "application/octet-stream"),
        ("png", "image/png"),
        ("svg", "image/svg+xml"),
        ("tif", "application/octet-stream"),
        ("wmf", "image/x-wmf")
    )
    for extension_contenttype in extension_contenttypes
        ext, ct = extension_contenttype
        # do not add the extension if it is already defined in the template
        isnothing(findfirst(x -> (x.name == "Default" && x["Extension"] == ext), elements(r))) || continue
        addelement!(r, "Default Extension=\"$ext\" ContentType=\"$ct\"")
    end
    zip_newfile(w, path; compress=true)
    prettyprint(w, doc)
    zip_commitfile(w)
end

# Support reading a template from file path or from pre-read file data.
function read_template(template_path::AbstractString)
    template_path = abspath(template_path)
    template_isfile = isfile(template_path)
    if !template_isfile
        error(
            "No file found at template path: $(repr(template_path))",
        )
    end
    read(template_path)
end
read_template(template_data::AbstractVector{UInt8}) = template_data

"""
```julia
Base.write(
    filepath::String,
    p::Presentation;
    overwrite::Bool=false,
    open_ppt::Bool=true,
    template_path::String="no-slides.pptx",
)
```

* `filepath::String` Desired presentation filepath.
* `pres::Presentation` Presentation object to write.
* `overwrite = false` argument for overwriting existing file.
* `open_ppt = true` open powerpoint after it is written.
* `template_path::String` path to an (empty) pptx that serves as template.

# Examples
```
julia> using PPTX

julia> slide = Slide()

julia> text = TextBox("Hello world!")

julia> push!(slide, text)

julia> pres = Presentation()

julia> push!(pres, slide)

julia> write("hello_world.pptx", pres)
```
"""
function Base.write(
    filepath::String,
    p::Presentation;
    overwrite::Bool=false,
    open_ppt::Bool=true,
    template_path::Union{String, Vector{UInt8}}=DEFAULT_TEMPLATE_DATA,
)
    template_reader = ZipBufferReader(read_template(template_path))

    if !endswith(filepath, ".pptx")
        filepath *= ".pptx"
    end

    filepath = abspath(filepath)
    filedir, filename = splitdir(filepath)

    mkpath(filedir)

    if !overwrite && isfile(filepath)
        error(
            "File $(repr(filepath)) already exists use \"overwrite = true\" or a different name to proceed",
        )
    end

    mktemp(filedir) do temp_path, temp_out
        ZipWriter(temp_out; own_io=true) do w
            update_presentation_state!(p, template_reader)
            update_slide_nrs!(p)
            write_relationships!(w, p)
            write_presentation!(w, p)
            write_slides!(w, p, template_reader)
            write_shapes!(w, p)
            update_table_style!(w, template_reader)
            add_contenttypes!(w, template_reader)
            # copy over any files from the template
            # but don't overwrite any files in w
            for i in zip_nentries(template_reader):-1:1
                local name = zip_name(template_reader, i)
                if !endswith(name,"/")
                    if !zip_name_collision(w, name)
                        local compress = zip_iscompressed(template_reader, i)
                        zip_data = zip_readentry(template_reader, i)
                        zip_newfile(w, name; compress)
                        write(w, zip_data)
                        zip_commitfile(w)
                    end
                end
            end
        end
        mv(temp_path, filepath; force=overwrite)
    end

    if open_ppt
        try
            DefaultApplication.open(filepath)
        catch err
            @warn "Could not open file $(repr(filepath))"
            bt = backtrace()
            print(sprint(showerror, err, bt))
        end
    end
    return nothing
end
