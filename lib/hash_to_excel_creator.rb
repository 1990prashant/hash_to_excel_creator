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
      @workbook.add_worksheet(name: "Sample") do |sheet|
        sheet.add_row(hash.keys.flatten.compact)
        10.times do 
          sheet.add_row(hash.values.flatten.compact)
        end
        sheet.add_row(["ends"])
      end
    elsif hash.is_a? Array

    end
    stream = @package.to_stream()
    File.open('lib/first.xlsx', 'w+') { |f| f.write(stream.read) }
  end

end