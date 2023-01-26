
#=
Most PPTX xml's contain the following attributes

For example in a 'slide'
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">

We use `AbstractDict{String,String}` to define attributes

=#
function main_attributes()::Vector{OrderedDict}
    return OrderedDict[
        OrderedDict("xmlns:a" => "http://schemas.openxmlformats.org/drawingml/2006/main"),
        OrderedDict(
            "xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        OrderedDict(
            "xmlns:p" => "http://schemas.openxmlformats.org/presentationml/2006/main"
        )
    ]
end

function get_slide_layout_path(layout::Int, unzipped_ppt_dir::String=".")
    filename = "slideLayout$layout.xml"
    path = joinpath(unzipped_ppt_dir, "slideLayouts", filename)
    return abspath(path)
end

function xpath_to_find_sp_type(type_name::String)
    return "//p:spTree/p:sp/p:nvSpPr[1]/p:nvPr[1]/p:ph[1][@type=\"$type_name\"][1]/ancestor::p:sp[1]"
end

function get_title_shape_node(s::Slide, unzipped_ppt_dir::String=".")
    get_title_shape_node(s.layout, unzipped_ppt_dir)
end

function get_title_shape_node(layout::Int, unzipped_ppt_dir::String=".")
    layout_path = get_slide_layout_path(layout, unzipped_ppt_dir)
    layout_doc = EzXML.readxml(layout_path)

    # the xpath way to find things
    # note: on layout2 and forth it's type="title", layout1 uses type="ctrTitle"
    xpath = xpath_to_find_sp_type("title")
    title_shape_node = findfirst(xpath, root(layout_doc))
    if isnothing(title_shape_node)
        xpath = xpath_to_find_sp_type("ctrTitle")
        title_shape_node = findfirst(xpath, root(layout_doc))
    end

    return title_shape_node
end

# <p:sp> input
function get_title_from_shape_node(title_shape_node::EzXML.Node)
    r = findfirst("./p:txBody[1]/a:p[1]/a:r[1]", title_shape_node)
    t = findfirst("./a:t", r)
    return t
end

# <p:sp> input
function update_xml_title!(title_shape_node::EzXML.Node, title::String)
    t = get_title_from_shape_node(title_shape_node)
    if isnothing(t)
        PPTX.link_node!(r, Dict("a:t" => PPTX.TextBody(title)))
    else
        t.content = title
    end

    return nothing
end

get_shape_ids(doc::EzXML.Document) = get_shape_ids(root(doc))

function get_shape_ids(node::EzXML.Node)
    # xpath to find something with an unregistered namespace must use local-name
    cNvPr_vector = findall("//*[local-name()='p:cNvPr' or local-name()='cNvPr']", node)
    ids = [parse(Int, el["id"]) for el in cNvPr_vector]
    return ids
end

# <p:sp> input
function update_shape_id!(sp_node::EzXML.Node, id::Int)
    cNvPr = findfirst("./p:nvSpPr/p:cNvPr", sp_node)
    cNvPr["id"] = "$id"
    return nothing
end

function has_empty_table_list(unzipped_ppt_dir::String=".")
    table_style_path = abspath(joinpath(unzipped_ppt_dir, "tableStyles.xml"))
    table_style_doc = EzXML.readxml(table_style_path)
    tblStyles = findall("//a:tblStyleLst/a:tblStyle", root(table_style_doc))
    return isnothing(tblStyles) || isempty(tblStyles)
end