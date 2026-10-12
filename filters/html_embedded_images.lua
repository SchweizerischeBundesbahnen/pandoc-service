-- Load the images app/html_embedded_images.py took out of the HTML into the
-- media bag, which the DOCX writer embeds them from.
--
-- pandoc runs with --sandbox and reads no file a document names. This filter
-- reads the images itself, and only from the directory the service gives it in
-- PANDOC_SERVICE_IMAGES_DIR, which a document cannot set. It reads only the
-- names the service gives an image. Their token is random, so a document naming
-- a file of its own gets nothing.
--
-- pandoc's HTML reader resolves a src against a <base href>, a data: URI aside,
-- so the name may arrive at the end of an address. The filter takes the name
-- from there and writes it back as the src, which the DOCX writer and the
-- {{IMG:...}} placeholders of filters/inline_styles.lua then carry.
--
-- It runs first: filters/inline_styles.lua turns some images into raw OOXML.

local directory = os.getenv("PANDOC_SERVICE_IMAGES_DIR")

local MIME_TYPES = {
  png = "image/png",
  jpg = "image/jpeg",
  gif = "image/gif",
  bmp = "image/bmp",
  tiff = "image/tiff",
}

local loaded = {}

-- The name of an extracted image at the end of src, else nil.
local function image_name(src)
  local name = src:match("([^/]+)$")
  if name and name:match("^pandoc%-service%-image%-%x+%-%d+%.%l+$") then
    return name
  end
  return nil
end

-- Load the image of a name into the media bag once; whether it is there.
local function load(name)
  if loaded[name] then
    return true
  end
  local mime_type = MIME_TYPES[name:match("%.(%l+)$")]
  if not mime_type then
    return false
  end
  local file = io.open(directory .. "/" .. name, "rb")
  if not file then
    return false
  end
  local contents = file:read("a")
  file:close()
  pandoc.mediabag.insert(name, mime_type, contents)
  loaded[name] = true
  return true
end

if not directory or directory == "" then
  return {}
end

return {
  {
    Image = function(image)
      local name = image_name(image.src)
      if name and load(name) then
        image.src = name
        return image
      end
    end,
  },
}
