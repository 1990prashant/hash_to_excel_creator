# Hash To Excel Creator

Hash To excel creator will conver your hash or array of hashes into excel with allowing to create multiple worksheet in a single workbook.

# Usage

Add follwing line into your Gemfile

gem 'hash_to_excel_creator', github: '1990prashant/hash_to_excel_creator'

Create a new object of HashToExcelCreator Class and call get_excel() method

hte = HashToExcelCreator.new
hte.get_excel(hash, file_name)

# Example hash

{
  sheet_name: "Name of your sheet"
  a: [1, 2], 
  b: [1, 3]
}