class HashToExcelCreator

  require 'axlsx'

  def initialize
    @package = Axlsx::Package.new
    @workbook = @package.workbook
  end

  def get_excel(hash = get_hash)
    convert_hash_to_excel(hash)
  end

  def get_hash
    h = {}
    ('a'..'z').to_a.each_with_index do |n, i|
      h[n] = i + 1
    end
    h
  end

  def convert_hash_to_excel(hash)
    if hash.is_a? Hash
      @workbook.add_worksheet(name: "Prashant") do |sheet|
        # add_heading(hash.keys, sheet)
        # add_values(hash.values, sheet)
        sheet.add_row(hash.keys.flatten)
        get_final_values(hash.values).each do |arr|
          sheet.add_row(arr.flatten)
        end
      end
    elsif hash.is_a? Array

    end
    stream = @package.to_stream()
    file_name = File.expand_path('../lib/sample.xlsx')
    File.open(file_name, 'w+') { |f| f.write(stream.read) }
  end

  def add_heading(headings, sheet)
    append_row(headings, sheet)
  end

  def add_values(values_arr, sheet)
    get_final_values(values_arr).each do |arr|
      append_row(arr, sheet)
    end
  end

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