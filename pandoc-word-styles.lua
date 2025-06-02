-- pandoc-word-styles.lua (More Compatible Version)
-- Original concept by: Cheuk Kin CHAU (ickc)
-- Adapted for robustness, to remove problematic dependencies, and for wider Pandoc compatibility.
-- Purpose: Converts CSS styles in span attributes to OpenXML for DOCX.

-- Standard color names to hex mapping
local color_names = {
  black = "000000", silver = "C0C0C0", gray = "808080", white = "FFFFFF",
  maroon = "800000", red = "FF0000", purple = "800080", fuchsia = "FF00FF",
  green = "008000", lime = "00FF00", olive = "808000", yellow = "FFFF00",
  navy = "000080", blue = "0000FF", teal = "008080", aqua = "00FFFF"
}

-- Helper function to safely escape XML attribute values
-- This is used if pandoc.utils.encode_xml_text (Pandoc 2.18+) is not available.
local function custom_escape_xml_attr_value(s_raw)
  local s = tostring(s_raw)
  s = s:gsub("&", "&amp;")
  s = s:gsub("<", "&lt;")
  s = s:gsub(">", "&gt;")
  s = s:gsub("\"", "&quot;")
  -- s = s:gsub("'", "&apos;") -- &apos; is not strictly needed for attributes quoted with "
  return s
end

-- Determine which XML escaping function to use
local active_encode_xml_text
if pandoc.utils.encode_xml_text then
  active_encode_xml_text = pandoc.utils.encode_xml_text
else
  active_encode_xml_text = custom_escape_xml_attr_value
end

-- Helper function to convert color names or hex strings to a 6-digit hex string
local function get_color_hex(color_str)
  if color_str == nil then return nil end
  local s = color_str:lower():gsub("#", "")

  if color_names[s] then
    return color_names[s]:upper()
  end

  if s:match("^[0-9a-f]+$") then
    if #s == 3 then
      s = s:gsub(".", "%1%1")
    end
    if #s == 6 then
      return s:upper()
    end
  end
  return nil
end

-- Parses the style string and returns a list of OpenXML elements for <w:rPr>
local function parse_style_to_rpr_children(style_str)
  local rpr_children = {}

  for property, value in style_str:gmatch("([%w-]+) *: *([^;]+);?") do
    local prop_lower = property:lower()
    local val_lower = value:lower()

    if prop_lower == "color" then
      local color_val = get_color_hex(value)
      if color_val then
        table.insert(rpr_children, {tag = "w:color", attributes = {["w:val"] = color_val}})
      end
    elseif prop_lower == "background-color" or prop_lower == "background" then
      local color_val = get_color_hex(value)
      if color_val then
        table.insert(rpr_children, {tag = "w:shd", attributes = {["w:val"] = "clear", ["w:color"] = "auto", ["w:fill"] = color_val}})
      end
    elseif prop_lower == "font-weight" and val_lower == "bold" then
      table.insert(rpr_children, {tag = "w:b"})
      table.insert(rpr_children, {tag = "w:bCs"})
    elseif prop_lower == "font-style" and val_lower == "italic" then
      table.insert(rpr_children, {tag = "w:i"})
      table.insert(rpr_children, {tag = "w:iCs"})
    elseif prop_lower == "text-decoration" and val_lower == "underline" then
      table.insert(rpr_children, {tag = "w:u", attributes = {["w:val"] = "single"}})
    end
  end
  return rpr_children
end

-- Main function to process Span elements
function Span(el)
  if el.attributes.style then
    local rpr_children = parse_style_to_rpr_children(el.attributes.style)

    if #rpr_children > 0 then
      local rpr_xml_parts = {"<w:rPr>"}
      for _, child_tag_info in ipairs(rpr_children) do
          local tag_name = child_tag_info.tag
          local attrs_str = ""
          if child_tag_info.attributes then
              for k, v_raw in pairs(child_tag_info.attributes) do
                  -- Use the determined XML escaping function
                  local v_escaped = active_encode_xml_text(tostring(v_raw)) -- This was line 82
                  attrs_str = attrs_str .. string.format(' %s="%s"', k, v_escaped)
              end
          end
          table.insert(rpr_xml_parts, string.format("<%s%s/>", tag_name, attrs_str))
      end
      table.insert(rpr_xml_parts, "</w:rPr>")
      local rpr_props_xml = table.concat(rpr_xml_parts)

      table.insert(el.content, 1, pandoc.RawInline('openxml', rpr_props_xml))
      el.attributes.style = nil
    end
  end
  return el
end

return {
  {Span = Span}
}