function link_node!(
    node::EzXML.Node, pair::Pair{<:Any,<:Union{AbstractDict,AbstractVector}}
)
    (key, value) = pair
    key_node = ElementNode(key)
    link_node!(key_node, value)
    link!(node, key_node)
end

function link_node!(node::EzXML.Node, pair::Pair{<:Any,TextBody})
    (key, value) = pair
    key_node = ElementNode(key)
    text_node = TextNode(String(value))
    link!(key_node, text_node)
    link!(node, key_node)
end

function link_node!(node::EzXML.Node, pair::Pair{<:Any,<:AbstractString})
    (key, value) = pair
    key_node = AttributeNode(key, value)
    link!(node, key_node)
end

function link_node!(node::EzXML.Node, pair::Pair{<:Any,<:Union{Nothing,Missing}})
    (key, value) = pair
    key_node = ElementNode(key)
    link!(node, key_node)
end

function link_node!(node::EzXML.Node, dict::AbstractDict)
    for pair in dict
        link_node!(node, pair)
    end
end

function link_node!(node::EzXML.Node, list::AbstractVector)
    for elt in list
        link_node!(node, elt)
    end
end

function link_node!(node::EzXML.Node, tup::NamedTuple)
    for (key, value) in Base.zip(keys(tup), values(tup))
        key_node = AttributeNode(key, value)
        link!(node, key_node)
    end
end

function xml_document(xml::AbstractDict)::EzXML.Document
    doc = EzXML.XMLDocument()
    link_node!(doc.node, xml)
    return doc
end