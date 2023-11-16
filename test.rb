require "google_drive"

# Initializes a session with credentials from config.json.
session = GoogleDrive::Session.from_config("config.json")

# Fetches the first worksheet from the specified spreadsheet by its key.
spreadsheet_key = "14Joknl8QvTxZHkoZH_3yNu471qhBg-81We5svCkEPbY"
ws = session.spreadsheet_by_key(spreadsheet_key).worksheets[0]


class Column
  include Enumerable

  def initialize(worksheet, col_index)
    @worksheet = worksheet
    @col_index = col_index
    @values = @worksheet.rows.drop(1).map { |row| row[@col_index] }
    reject_unwanted_values!
  end

  def each(&block)
    @values.each(&block)
  end

  def [](index)
    @values[index]
  end

  
  def []=(index, value)
    
    row_index = index + 2
    
    @values[index] = value
    
    @worksheet[row_index, @col_index + 1] = value
    @worksheet.save
  end

  def inspect
    @values.inspect
  end

  private

  
  def reject_unwanted_values!
    @values.reject! do |cell|
      cell_value = cell.to_s.strip.downcase
      cell_value.empty? || cell_value == 'total' || cell_value == 'subtotal'
    end
  end
end

class Table
  include Enumerable

  attr_reader :headers, :worksheet

  def initialize(worksheet)
    @worksheet = worksheet
    @headers = load_headers
  end

  def load_headers
    headers = {}
    @worksheet.rows[0].each_with_index do |header, index|
      normalized_header = normalize_header(header)
      headers[normalized_header] = index + 1
    end
    headers
  end

  def row(row_index)
    
    actual_row_index = row_index + 1  
    row = @worksheet.rows[actual_row_index - 1]  
    return nil if row_contains_keywords?(row) || row_empty?(row)
    row
  end

  def each
    @worksheet.rows.drop(1).each do |row|
      next if row_contains_keywords?(row) || row_empty?(row)
      yield row.map { |cell| cell.to_s.strip }  
    end
  end

  def [](header_name)
    col_index = @headers[normalize_header(header_name)]
    raise ArgumentError, "Header not found" unless col_index
    Column.new(@worksheet, col_index - 1)
  end

  def method_missing(method_name, *arguments, &block)
    method_name_str = method_name.to_s
    normalized_method_name = normalize_header(method_name_str)
  
    if @headers.has_key?(normalized_method_name)
      col_index = @headers[normalized_method_name] - 1
      if arguments.empty?
        
        Column.new(@worksheet, col_index)
      else
        
        value = arguments.first
        row = @worksheet.rows.drop(1).find { |r| r[col_index].to_s == value.to_s }
        row unless row_contains_keywords?(row) || row_empty?(row)
      end
    elsif normalized_method_name.end_with?('_sum') || normalized_method_name.end_with?('_avg')
      handle_aggregation(normalized_method_name)
    else
      super
    end
  end
  

  private

  def handle_aggregation(normalized_method_name)
    column_name = normalized_method_name.sub(/_(sum|avg)$/, '')
    column_values = self[column_name].reject { |val| val.to_s.downcase.match(/total|subtotal/) }.map(&:to_f)
    if normalized_method_name.end_with?('_sum')
      column_values.sum
    elsif normalized_method_name.end_with?('_avg')
      column_values.empty? ? 0 : column_values.sum / column_values.size
    end
  end

  def row_contains_keywords?(row)
    row.any? { |cell| cell.to_s.downcase.match(/total|subtotal/) }
  end

  def row_empty?(row)
    row.all?(&:nil?) || row.reject(&:nil?).all? { |cell| cell.to_s.strip.empty? }
  end

  def respond_to_missing?(method_name, include_private = false)
    normalized_method_name = normalize_header(method_name.to_s)
    @headers.has_key?(normalized_method_name) || super
  end

  def normalize_header(header_name)
    header_name.strip.downcase.gsub(/\s+/, '_')
  end
end


table = Table.new(ws)

table["Treca Kolona"][3] = 2000

puts "Vrednosti vece od 10 u 'Prva kolona':"
puts table.prva_kolona.select { |cell| cell.to_i > 10 }.inspect


puts "Proizvod svih vrednosti u 'Prva kolona':"
puts table.prva_kolona.reduce(1) { |product, cell| product * cell.to_i }.inspect


puts "Povecane vrednosti u 'Prva kolona':"
puts table.prva_kolona.map { |cell| cell.to_i + 1 }.inspect

puts "Red sa vrednosti '12' u 'treca_kolona':"
puts table.treca_kolona('12').inspect

puts "Suma: 'Treca kolona':"
puts table.treca_kolona_sum.inspect

puts "Prosecna vrednost 'Peta kolona':"
puts table.peta_kolona_avg.inspect


puts "Vrednosti u 'Cetvrta kolona':"
puts table.cetvrta_kolona.inspect


puts "Prvi red vrednosti:"
puts table.row(1).inspect


puts "Sve celije vrednosti:"
table.each { |cell| puts cell }


puts "Vrednosti u 'Prva kolona':"
puts table["Prva kolona"].inspect


puts "Cetvrti element u 'Peta kolona':"
puts table["Peta kolona"][4].inspect


def get_table_as_array(worksheet)
  rows = worksheet.num_rows
  cols = worksheet.num_cols
  (1..rows).map do |row|
    (1..cols).map { |col| worksheet[row, col] }
  end
end


table_values = get_table_as_array(ws)
puts "Vrednosti tabele kao niz:"
puts table_values.inspect


  
  
