require "#{Docsplit::ROOT}/vendor/jodconverter/jodconverter-core-3.0-beta-4.jar"

module Docsplit
  class JavaConverter
    at_exit { JavaConverter.stop }

    @@lock = Mutex.new

    def self.convert(source, destination)
      converter.convert(java.io.File.new(source), java.io.File.new(destination))
    end

    def self.start
      converter && true
    end

    def self.stop
      @@lock.synchronize do
        if @@manager
          @@manager.stop
          @@manager = nil
        end
        @@converter = nil
      end
      true
    end

    private

    def self.converter
      @@lock.synchronize do
        unless @@converter
          config = org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration.new
          config.setOfficeHome(Docsplit::OFFICE_HOME) if Docsplit::OFFICE_HOME.length > 0
          @@manager = config.buildOfficeManager
          @@manager.start

          format_registry =
            org.artofsolving.jodconverter.document.JsonDocumentFormatRegistry.new(
              File.read("#{Docsplit::ROOT}/vendor/conf/document-formats.js")
          )
          @@converter = org.artofsolving.jodconverter.OfficeDocumentConverter.new(@@manager, format_registry)
        end
      end
      @@converter
    end
  end
end
