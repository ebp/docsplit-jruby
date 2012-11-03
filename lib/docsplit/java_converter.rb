require "#{Docsplit::ROOT}/vendor/jodconverter/jodconverter-core-3.0-beta-4.jar"

module Docsplit
  class JavaConverter
    at_exit { JavaConverter.stop }

    @lock = Mutex.new

    class << self

      def convert(source, destination)
        converter = org.artofsolving.jodconverter.OfficeDocumentConverter.new(manager, format_registry)
        converter.convert(java.io.File.new(source), java.io.File.new(destination))
      end

      def stop
        @lock.synchronize do
          if @manager
            @manager.stop
            @manager = nil
          end
        end
        true
      end

      private

      def manager
        @lock.synchronize do
          @manager ||= begin
                         config = org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration.new
                         config.setOfficeHome(Docsplit::OFFICE) if Docsplit::OFFICE.length > 0
                         mgr = config.buildOfficeManager
                         mgr.start
                         mgr
                       end
        end
      end

      def format_registry
        @lock.synchronize do
          @format_registry ||= org.artofsolving.jodconverter.document.JsonDocumentFormatRegistry.new(File.read("#{Docsplit::ROOT}/vendor/conf/document-formats.js"))
        end
      end
    end
  end
end
