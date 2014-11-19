require './win32ole-ext'

module Excel
	def Excel.new(visible = true, displayAlerts = false)
	    excel = WIN32OLE.new_with_const('Excel.Application', Excel)
    	excel.visible = visible
    	excel.displayAlerts = displayAlerts
    	return excel
  	end
	def Excel.runDuring(visible = true, displayAlerts = false, &block)
		begin
			excel = new(visible, displayAlerts)
			block.call(excel)
		ensure
			excel.quit
		end
		excel = WIN32OLE.new_with_const('Excel.Application', Excel)
		excel.visible = visible
		excel.displayAlerts = displayAlerts
		return excel
	end

	module CellAccessor
		def initialize(range)
			@range = range
			@cell_cache = Hash.new
		end

		def lookup_cell_named(name)
			return @range.range(name)
		end

		def cell_named(name)
			cell = @cell_cache[name]
			if cell == nil then
				cell = lookup_cell_named(name)
				@cell_cache[name] = cell
			end
			return cell
		end

		def []=(name, value)
			cell = cell_named(name)
			cell.value = vale
		end

		def [](name)
			cell = cell_named(name)
			return cell.value
		end
	end
end
