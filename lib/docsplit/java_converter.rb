require "#{Docsplit::ROOT}/vendor/jodconverter/jodconverter-core-3.0-beta-4.jar"

module Docsplit
  class JavaConverter
    at_exit { JavaConverter.stop }

    def self.start
      stop if @manager
      @manager = manager_config.buildOfficeManager
      @manager.start

      @manager
    end

    def self.convert(source, destination)
      @manager ||= self.start

      converter = org.artofsolving.jodconverter.OfficeDocumentConverter.new(@manager, format_registry)
      converter.convert(java.io.File.new(source), java.io.File.new(destination))
    end

    def self.stop
      if @manager
        @manager.stop
        @manager = nil
      end

      true
    end

    private

    def self.format_registry
      @format_registry ||= org.artofsolving.jodconverter.document.JsonDocumentFormatRegistry.new(File.read("#{Docsplit::ROOT}/vendor/conf/document-formats.js"))
    end

    def self.manager_config
      @manager_config ||= begin
                            config = org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration.new
                            config.setOfficeHome(Docsplit::OFFICE) if Docsplit::OFFICE.length > 0
                            config
                          end
    end
  end
end
