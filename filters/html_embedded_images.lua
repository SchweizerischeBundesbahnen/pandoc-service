-- Load the images app/html_embedded_images.py took out of the HTML into the
-- media bag, which the DOCX writer embeds them from.
--
-- pandoc runs with --sandbox and reads no file a document names. This filter
-- reads the images itself, and only from the directory the service gives it in
-- PANDOC_SERVICE_IMAGES_DIR, which a document cannot set. It reads only the
-- names the service gives an image. Their token is random, so a document naming
-- a file of its own gets nothing.
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

local function load(src)
  if loaded[src] then
    return
  end
  local extension = src:match("^pandoc%-service%-image%-%x+%-%d+%.(%l+)$")
  local mime_type = extension and MIME_TYPES[extension]
  if not mime_type then
    return
  end
  local file = io.open(directory .. "/" .. src, "rb")
  if not file then
    return
  end
  local contents = file:read("a")
  file:close()
  pandoc.mediabag.insert(src, mime_type, contents)
  loaded[src] = true
end

if not directory or directory == "" then
  return {}
end

return {
  {
    Image = function(image)
      load(image.src)
    end,
  },
}
