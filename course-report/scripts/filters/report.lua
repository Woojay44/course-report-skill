-- pandoc Lua 过滤器：标题编号、图表编号（需要 pandoc >= 3.0）
-- 用法：pandoc ... --lua-filter report.lua -M numbering=cn|decimal|manual
-- 编号直接写进标题/题注文字，这样目录里的编号和正文完全一致，两个后端通用。

local scheme = "cn"
local counters = {0, 0, 0, 0}
local fig_n, tab_n = 0, 0

local unnumbered = {
  ["参考文献"] = true, ["参考资料"] = true, ["致谢"] = true, ["谢辞"] = true,
  ["references"] = true, ["acknowledgements"] = true, ["acknowledgments"] = true,
}

local cn_digits = {"一", "二", "三", "四", "五", "六", "七", "八", "九"}

local function cn_num(n)
  if n < 10 then
    return cn_digits[n]
  elseif n < 20 then
    return "十" .. (n > 10 and cn_digits[n - 10] or "")
  elseif n < 100 then
    local tens, units = n // 10, n % 10
    return cn_digits[tens] .. "十" .. (units > 0 and cn_digits[units] or "")
  end
  return tostring(n)
end

local function read_meta(meta)
  if meta.numbering then
    scheme = pandoc.utils.stringify(meta.numbering)
    -- 读完就删：避免未知元数据流进 pandoc 的输出
    meta.numbering = nil
    return meta
  end
end

local function number_header(h)
  if scheme == "manual" or scheme == "none" then return nil end
  local text = pandoc.utils.stringify(h.content)
  if h.classes:includes("unnumbered") or unnumbered[text] or unnumbered[text:lower()] then
    return nil
  end
  local lvl = h.level
  if lvl > 4 then return nil end
  counters[lvl] = counters[lvl] + 1
  for i = lvl + 1, 4 do counters[i] = 0 end
  local prefix
  if scheme == "decimal" then
    local parts = {}
    for i = 1, lvl do parts[#parts + 1] = tostring(counters[i]) end
    prefix = table.concat(parts, ".") .. " "
  else
    if lvl == 1 then
      prefix = cn_num(counters[1]) .. "、"
    elseif lvl == 2 then
      prefix = "（" .. cn_num(counters[2]) .. "）"
    elseif lvl == 3 then
      prefix = tostring(counters[3]) .. ". "
    else
      prefix = "（" .. tostring(counters[4]) .. "）"
    end
  end
  h.content:insert(1, pandoc.Str(prefix))
  return h
end

local function prefix_caption(caption, label, n)
  local long = caption.long
  if #long == 0 then return false end
  local first = long[1]
  if first.content == nil then return false end
  first.content:insert(1, pandoc.Str(label .. n .. "　"))
  return true
end

local function number_table(t)
  if #t.caption.long == 0 then return nil end
  tab_n = tab_n + 1
  prefix_caption(t.caption, "表", tab_n)
  return t
end

local function number_figure(f)
  if #f.caption.long == 0 then return nil end
  fig_n = fig_n + 1
  prefix_caption(f.caption, "图", fig_n)
  return f
end

-- 正文引用 [1]、[2][3] → 上标角标（参考文献节内的条目编号不动）。只认单个数字一对括号，
-- 多篇连引写成 [4][5]，不用 [4,5] / [1-3]。
-- 紧跟在英文字母/数字后的 [0] 之类（如 a[0]）视为下标写法，不转换。
local CITE = "%[%d+%]"

local function split_cites(s)
  local out, pos = {}, 1
  while true do
    local a, b = s:find(CITE, pos)
    while a and a > 1 and s:sub(a - 1, a - 1):match("%w") do
      a, b = s:find(CITE, b + 1)
    end
    if not a then break end
    -- 连续的 [2][3] 合成一个角标
    while true do
      local c, d = s:find(CITE, b + 1)
      if c ~= b + 1 then break end
      b = d
    end
    if a > pos then out[#out + 1] = pandoc.Str(s:sub(pos, a - 1)) end
    out[#out + 1] = pandoc.Superscript({pandoc.Str(s:sub(a, b))})
    pos = b + 1
  end
  if #out == 0 then return nil end
  if pos <= #s then out[#out + 1] = pandoc.Str(s:sub(pos)) end
  return out
end

local function cite_inlines(inlines)
  local res, changed = pandoc.Inlines({}), false
  for _, el in ipairs(inlines) do
    local parts = el.t == "Str" and split_cites(el.text) or nil
    if parts then
      changed = true
      for _, p in ipairs(parts) do
        -- 角标前的空格去掉：“重叠 [2]” → “重叠²”
        if p.t == "Superscript" and #res > 0 and res[#res].t == "Space" then
          res:remove(#res)
        end
        res:insert(p)
      end
    else
      res:insert(el)
    end
  end
  return changed and res or nil
end

local function superscript_cites(doc)
  local in_refs = false
  for i, blk in ipairs(doc.blocks) do
    if blk.t == "Header" and blk.level == 1 then
      local text = pandoc.utils.stringify(blk.content)
      in_refs = unnumbered[text] or unnumbered[text:lower()] or false
      in_refs = in_refs and text ~= "致谢" and text ~= "谢辞"
        and text:lower() ~= "acknowledgements" and text:lower() ~= "acknowledgments"
    elseif not in_refs then
      doc.blocks[i] = pandoc.walk_block(blk, { Inlines = cite_inlines })
    end
  end
  return doc
end

return {
  { Meta = read_meta },
  { Pandoc = superscript_cites },
  { Header = number_header, Table = number_table, Figure = number_figure },
}
