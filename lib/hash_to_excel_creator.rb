class HashToExcelCreator

  require 'axlsx'

  def initialize
    @package = Axlsx::Package.new
    @workbook = @package.workbook
  end

  def get_excel(hash = get_hash, file_name = "sample")
    convert_hash_to_excel(hash, file_name)
  end

  def get_hash
    h = {}
    ('a'..'z').to_a.each_with_index do |n, i|
      h[n] = i + 1
    end
    h
  end

  def convert_hash_to_excel(hash, user_file_name)
    if hash.is_a? Hash
      create_excel(hash)
    elsif hash.is_a? Array
      hash.each_with_index do |obj, index|
        create_excel(convert_to_symbolize_hash(obj), (index + 1))
      end
    end
    stream = @package.to_stream()
    file_name = File.expand_path("../#{user_file_name}.xlsx")
    File.open(file_name, 'w+') { |f| f.write(stream.read) }
  end

  def convert_to_symbolize_hash(obj)
    hash = {}
    obj.map{|key, value| hash[key.to_sym] = value}
    return hash
  end

  def create_excel(hash, sheet_number = 1)

    sheet_name = hash.has_key?(:sheet_name) ? hash[:sheet_name] : "Sheet-#{sheet_number}"
    @workbook.add_worksheet(name: sheet_name) do |sheet|
      sheet.add_row(hash.except(:sheet_name).keys.flatten)
      get_final_values(hash.except(:sheet_name).values).each do |arr|
        sheet.add_row(arr.flatten)
      end
    end
  end

  # def add_heading(headings, sheet)
  #   append_row(headings, sheet)
  # end

  # def add_values(values_arr, sheet)
  #   get_final_values(values_arr).each do |arr|
  #     append_row(arr, sheet)
  #   end
  # end

  def get_final_values(values_arr)
    result_arr = []
    (0..values_arr[0].size - 1).to_a.each do |i|
      temp = []
      (0..values_arr.size - 1).to_a.each do |index|
        temp << values_arr[index][i]
      end
      result_arr << temp
    end
    result_arr
  end

  def append_row(values, sheet)
    sheet.add_row(values.flatten)
  end

end