#!/usr/bin/env ruby

require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'nokogiri'

class WordXmlManipulate
  def self.open(path, &block)
    self.new(path, &block)
  end

  def initialize(path, &block)
    @replace = {}
    if block_given?
      @zip = Zip::ZipFile.open(path)
      yield(self)
      @zip.close
    else
      @zip = Zip::ZipFile.open(path)
    end
  end

  def merge(rec)
    puts "DEBUG: entering merge function"
    xml = @zip.read("word/document.xml")
    puts "DEBUG: finished reading document"
    doc = Nokogiri::XML(xml) {|x| x.noent}
    puts "DEBUG: finished parsing document with nokogiri"
    (doc/"//w:p").each do |field|
      text_nodeset = (field/".//w:t").first
      if text_nodeset
        if (rec[text_nodeset.inner_html])
          text_nodeset.inner_html = rec[text_nodeset.inner_html].to_s
        end
      end
    end
    @replace["word/document.xml"] = doc.serialize :save_with => 0
  end

  def save(path)
    Zip::ZipFile.open(path, Zip::ZipFile::CREATE) do |output|
      @zip.each do |entry|
        output.get_output_stream(entry.name) do |o|
          if @replace[entry.name]
            o.write(@replace[entry.name])
          else
            o.write(@zip.read(entry.name))
          end
        end
      end
    end
    @zip.close
  end
end

if __FILE__ == $0
  file = ARGV[0]
  working_file = ARGV[1] || file.sub(/\.docx/, '-merged.docx')
  w = WordXmlManipulate.open(file) 
  w.merge('REPLACE_THIS' => 'Some new text', 'First_Name' => 'James', 'Last_Name' => 'Vee')
  w.save(working_file)
  puts"DEBUG: complete"
end