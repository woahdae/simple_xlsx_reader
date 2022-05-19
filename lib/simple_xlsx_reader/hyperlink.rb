# frozen_string_literal: true

module SimpleXlsxReader
  # We support hyperlinks as a "type" even though they're technically
  # represented either as a function or an external reference in the xlsx spec.
  #
  # Since having hyperlink data in our sheet usually means we might want to do
  # something primarily with the URL (store it in the database, download it, etc),
  # we go through extra effort to parse the function or follow the reference
  # to represent the hyperlink primarily as a URL. However, maybe we do want
  # the hyperlink "friendly name" part (as MS calls it), so here we've subclassed
  # string to tack on the friendly name. This means 80% of us that just want
  # the URL value will have to do nothing extra, but the 20% that might want the
  # friendly name can access it.
  #
  # Note, by default, the value we would get by just asking the cell would
  # be the "friendly name" and *not* the URL, which is tucked away in the
  # function definition or a separate "relationships" meta-document.
  #
  # See MS documentation on the HYPERLINK function for some background:
  # https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f
  class Hyperlink < String
    attr_reader :friendly_name

    def initialize(url, friendly_name = nil)
      @friendly_name = friendly_name
      super(url)
    end
  end
end
