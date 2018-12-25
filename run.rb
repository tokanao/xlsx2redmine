require 'rubyXL'
require 'rexml/document'
require 'uri'
require 'net/https'
require 'fileutils'
require 'date'
require 'pry'

# ref. http://www.redmine.org/projects/redmine/wiki/Rest_api

# INTERNAL ERROR!!! invalid byte sequence in Windows-31J
Encoding.default_external = 'UTF-8'

# 個人設定画面に表示されているAPIキー
@api_key = '{your Redmine API key}'
@url = '{your Redmine URL}'
@filepath = './issues_sample.xlsx'

# 対象シートを限定する
# table_names = %w(issue project user)
table_names = %w(issue)

# 読込開始行(0 origin)
@row_start_idx = 2


@isdebug = true if ARGV[0] == "-d"

@book = RubyXL::Parser.parse(@filepath)


# 関連チケット情報とカスタム項目の情報収集
def create_template_xml(sheet)
  xml = REXML::Document.new
  table_name = sheet.sheet_name
  xml_root = xml.add_element(table_name)
  template_xmls = {table_name=>xml_root}

  header_cols = sheet[0]
  header_cols.size.times do |i|

    break if header_cols[i].nil?

    elm_val = header_cols[i].value.to_s

    case table_name
    when "issue"
      # フォーマットチェック e.g."custom_field:1"
      if /\d/ =~ elm_val then
        # マッチしたテキストの手前の文字列を取得 - "custom_field"
        elm_val = elm_val.split(":")[0]
        flds = xml_root.add_element("custom_fields")
        flds.add_attribute("type", "array")
        element = flds.add_element(elm_val)

        # マッチした文字列取得 - ":1"
        id = Regexp.last_match(0)
        element.add_attribute("id", id)
      else
        # IssueRelation
        if %w(issue_from_id issue_to_id relation_type).include?(elm_val) then
          relation_root = template_xmls["relation"]
          if relation_root.nil?
            relation_template = REXML::Document.new
            relation_root = relation_template.add_element("relation")
          end
          relation_root.add_element(elm_val)
          template_xmls["relation"] = relation_root
        else
          xml_root.add_element(elm_val)
        end
      end
    else
        xml_root.add_element(elm_val)
    end
  end
  template_xmls
end

# 送信用 xml 作成
def create_xml_data(sheet, template_xml)
  data_xmls = Array::new
  (sheet.count - @row_start_idx).times do |i|

    # Skip record
    cols = sheet[i + @row_start_idx]

    # EOF by first column
    break if cols[0].nil? 
    
    data_xmls[i] = template_xml.deep_clone()

    cols.size.times do |j|

      title_cell = sheet[0][j]
      break if title_cell.nil?

      if !cols[j].nil? 

        cell = cols[j].value 
        text = cell.to_s
        if cell.kind_of?(DateTime) then
          text = cell.strftime("%Y-%m-%d")
        end

        # カスタムフィールドはセルタイトル(custom_fields:9)とタグ名が(custom_fields)が異なるため、分岐
        title_value = title_cell.value
        if title_value.include?("custom_field") then
          element = data_xmls[i].root.elements["custom_fields"]
          next if element.nil? 
          element[0].add_element("value").add_text text
        else
          element = data_xmls[i].root.elements[title_value]
          # 項目値が"0"になる場合は登録内容から除外する
          next if element.nil? or text == "0"
          element.add_text text
        end
      end
    end
  end
  data_xmls
end

# Redmine へリクエスト送信
def send_request(sheet, template_xml, data_xmls)

  update_flg = false

  uri = URI.parse(@url)
  http = Net::HTTP.new(uri.host, uri.port) 

  # SSL
  http.use_ssl = true
  http.verify_mode = OpenSSL::SSL::VERIFY_NONE

  http.set_debug_output $stderr if @isdebug

  http.start do |h|

    target_tables = template_xml.root.name + "s"
    p "---- #{target_tables} ----"

    data_xmls.size.times do |i|
      xml_root = data_xmls[i].root

      # EOF by first element
      first_element = xml_root[0]
      if target_tables == "relations"
        next if first_element.nil? || first_element.text.nil? 

        id = xml_root.elements["issue_from_id"].text
        uri = URI.parse(@url + "issues/#{id}/relations.xml")
        request = Net::HTTP::Post.new(uri.path)

      else
        break if first_element.nil? or first_element.text.empty? 

        # Decide new or update
        id = xml_root.elements["id"].text
        if id.nil? or id.empty? then 
          uri = URI.parse(@url + "#{target_tables}.xml")
          request = Net::HTTP::Post.new(uri.path)
        else
          uri = URI.parse(@url + "#{target_tables}/#{id}.xml")
          request = Net::HTTP::Put.new(uri.path)
        end
      end

      request.set_content_type("text/xml")
      request["X-Redmine-API-Key"] = @api_key

      request.body = xml_root.to_s
      response = h.request(request)

      # Check status 200:successs or 201:created
      if !%w(200 201).include?(response.code) then
        if @isdebug
          p request
          p request.body
          p response
          print response.body.encode("Shift_JIS", "UTF-8")
          debugger 
        else
          p "#{i + 1}件目... response.code = #{response.code}"
        end
        next
      end

      # Update key id
      if id.nil? or id.empty?
        update_flg = true

        res_xml = REXML::Document.new(response.body)
        id = res_xml.root.elements["id"].text

        col_idx = xml_root.index(xml_root.elements["id"])
        sheet.add_cell(i + @row_start_idx, col_idx, id)
      end

      p "#{i + 1}件目... ID = #{id}"

    end
  end

  if update_flg 
    @book.write(@filepath)
  end
end

def backup(file)
  name, ext = /\A(.+?)((?:\.[^.]+)?)\z/.match(file, &:captures)
  dest = name + "_#{DateTime.now.strftime('%Y%m%d%H%M%S')}" + ext
  FileUtils.cp(file, dest)
end

# ------------------------------------------------------------------------------

backup(@filepath)

@book.sheets.size.times do |i|
  if table_names.include?(@book[i].sheet_name)
    sheet = @book[i]

    template_xmls = create_template_xml(sheet)
    template_xmls.each do |table, xml|
      data_xmls = create_xml_data(sheet, xml)
      send_request(sheet, xml, data_xmls)
    end
  end
end

exit



# vim:ts=2:sw=2:et
