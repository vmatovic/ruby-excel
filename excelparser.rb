require 'roo'
require 'spreadsheet'

class Column < Array

	attr_accessor :matrix
	attr_reader :matrix
	
	def initialize(col, matr)
		super(col)
		@matrix = matr
	end

	def sum
		sol = 0
		self.each do |el|
			if el.nil?
				sol += 0
			else
				sol += el
			end
		end
		sol
	end

	def method_missing(key, *args)
		text = key.to_s
		self.each_with_index do |s, i|
			s = s.downcase
			if s == text
				return @matrix[i+1]
			end
		end
		return 0
	end
end

class Reader
	include Enumerable
	attr_accessor :matrix, :xlsx, :cols
	attr_reader :matrix, :xlsx, :cols
	
	def initialize(input, sheet)
		@matrix = []
		@cols = Hash.new
		
		if input.include? ".xlsx"
			@xlsx = Roo::Spreadsheet.open(input, {:expand_merged_ranges => true})
			for row in @xlsx.sheet(sheet)
				ignore = false
				if row.count(nil) == row.length
					ignore = true
				else
					for el in row
						if el == "total" or el == "subtotal"
							ignore = true
						end
					end
				end
				if not ignore
					@matrix = @matrix + [row]
				end
			end
			
			first_ind = @xlsx.sheet(sheet).first_column
			last_ind = @xlsx.sheet(sheet).last_column
			for i in first_ind..last_ind
				col = @xlsx.sheet(sheet).column(i)
				col = col.compact
				col.reject! {|c| c == "total"}
				col.reject! {|c| c == "subtotal"}
				cols[col[0]] = Column.new(col[1..-1], @matrix)
			end
		elsif input.include? ".xls"
			@xlsx = Spreadsheet.open(input)
			for row in @xlsx.worksheet(sheet)
				ignore = false
				if row.count(nil) == row.length
					ignore = true
				else
					for el in row
						if el == "total" or el == "subtotal"
							ignore = true
						end
					end
				end
				if not ignore
					@matrix = @matrix + [row]
				end

			end
			for i in 0..@xlsx.worksheet(sheet).column_count
				tmp = @xlsx.worksheet(sheet).column(i)
				col = Array.new
				tmp.each do |c|
					col << c
				end
				cols[col[0]] = Column.new(col[1..-1], @matrix)
			end
		else
			abort "Pogresan naziv fajla!"
		end
		
		

	end
	
	def write_matrix
		for row in @matrix
			print row, "\n"
		end
		
	end
	
	def row(i)
		@matrix[i]
	end
	
	def each &block
		@matrix.each{|el| for e in el
							block.call(e)
							end}
	end
	
	def [](key)
		cols[key]
	end
	
	def method_missing(key, *args)
		text = key.to_s
		@cols.each do |key, value|
			k = key.gsub(" ", "").downcase
			if k == text
				return value
		
			end
		end
	end

end

puts 'ispisi ime fajla:'
input = gets.chomp
puts 'ispisi ime stranice:'
sheet = gets.chomp
excel_reader = Reader.new(input, sheet)
excel_reader.write_matrix()

puts excel_reader.row(0)[0]

excel_reader.each {|x| print x, ' '}
puts ''

puts 'Prva kolona:'
puts excel_reader["Prva kolona"]
print "Drugi el. prve kolone ", excel_reader["Prva kolona"][1]
puts ''
puts 'Druga kolona:'
puts excel_reader.drugakolona
puts excel_reader.trecakolona.sum
#excel_reader.prvakolona.rn2310

#print excel_reader.indeks.rn4918
