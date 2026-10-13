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

return {
  { Meta = read_meta },
  { Header = number_header, Table = number_table, Figure = number_figure },
}
