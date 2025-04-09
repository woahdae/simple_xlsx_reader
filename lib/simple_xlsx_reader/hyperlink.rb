# frozen_string_literal: true

module SimpleXlsxReader
  # We support hyperlinks as a "type" even though they're technically
  # represented either as a function or an external reference in the xlsx spec.
  #
  # In practice, hyperlinks are usually a link or a mailto. In the case of a
  # link, we probably want to follow it to download something, but in the case
  # of an email, we probably just want the email and not the mailto. So we
  # represent a hyperlink primarily as it is seen by the user, following the
  # principle of least surprise, but the url is accessible via #url.
  #
  # Microsoft calls the visible part of a hyperlink cell the "friendly name,"
  # so we expose that as a method too, in case you want to be explicit about
  # how you're accessing it.
  #
  # See MS documentation on the HYPERLINK function for some background:
  # https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f
  class Hyperlink < String
    attr_reader :friendly_name
    attr_reader :url

    def initialize(url, friendly_name = nil)
      @url = url
      @friendly_name = friendly_name&.to_s
      super(@friendly_name || @url)
    end
  end
end
