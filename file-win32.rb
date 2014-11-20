require 'win32ole'

module FileSystemObject
	@instance = nil
	def FileSystemObject.instance
		unless @instance then
			@instance = WIN32OLE.new('Scripting.FileSystemObject')
		end
		return @instance
	end
end