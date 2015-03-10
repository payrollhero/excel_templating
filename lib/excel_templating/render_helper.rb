module ExcelTemplating
  module RenderHelper
    class << self
      def letter_for(one_based_index)
        letters[one_based_index-1]
      end

      private
      def letters
        @row_letters ||= ('A' .. 'ZZ').to_a
      end
    end
  end
end
