# coding: utf-8
require 'i18n/tasks/reports/base'
require 'fileutils'

module I18n::Tasks::Reports
  class Spreadsheet < Base

    def save_report(path, opts)
      path = path.presence || 'tmp/i18n-report.xlsx'
      p = Axlsx::Package.new
      add_missing_sheet p.workbook
      add_unused_sheet p.workbook
      add_eq_base_sheet p.workbook
      p.use_shared_strings = true
      FileUtils.mkpath(File.dirname(path))
      p.serialize(path)
      $stderr.puts Term::ANSIColor.green "Saved to #{path}"
    end

    def frontend_report(path, opts)
      target_locale = opts[:locales].present? ?  opts[:locales].first : 'en'
      # TODO: make the frontend reference language a parameter, now using :it
      base_locale = 'it'

      path = path.presence || "tmp/i18n-report.#{target_locale}.xlsx"
      p = Axlsx::Package.new

      tree = task.missing_keys({locales: [target_locale], base_locale: base_locale})
      if tree.keys.any?
        $stderr.puts "#{tree.keys.count} missing keys"
        add_frontend_sheet p.workbook, tree
        p.use_shared_strings = true
        FileUtils.mkpath(File.dirname(path))
        p.serialize(path)
        $stderr.puts Term::ANSIColor.green "Report saved to #{path}"
      else
        $stderr.puts Term::ANSIColor.green "No missing keys"
      end
    end

    private

    def add_frontend_sheet(wb, tree)
      wb.styles do |s|
        type_cell = s.add_style :alignment => {:horizontal => :center}
        locale_cell  = s.add_style :alignment => {:horizontal => :center}
        regular_style = s.add_style
        wb.add_worksheet(name: missing_title(tree)) { |sheet|
          sheet.page_setup.fit_to :width => 1
          sheet.add_row [I18n.t('i18n_tasks.common.key'), 'English', 'Translation']

          style_header sheet
          tree.keys do |key, node|
            locale, type = node.root.data[:locale], node.data[:type]

            if type == :missing_diff
              sheet.add_row [key, task.t(key)],
              styles: [type_cell, regular_style, regular_style]
            end
          end
        }
      end
    end

    def add_missing_sheet(wb)
      tree = task.missing_keys
      wb.styles do |s|
        type_cell = s.add_style :alignment => {:horizontal => :center}
        locale_cell  = s.add_style :alignment => {:horizontal => :center}
        regular_style = s.add_style
        wb.add_worksheet(name: missing_title(tree)) { |sheet|
          sheet.page_setup.fit_to :width => 1
          sheet.add_row [I18n.t('i18n_tasks.common.type'), I18n.t('i18n_tasks.common.locale'), I18n.t('i18n_tasks.common.key'), I18n.t('i18n_tasks.common.base_value')]
          style_header sheet
          tree.keys do |key, node|
            locale, type = node.root.data[:locale], node.data[:type]
            sheet.add_row [missing_type_info(type)[:summary], locale, key, task.t(key)],
            styles: [type_cell, locale_cell, regular_style, regular_style]
          end
        }
      end
    end

    def add_eq_base_sheet(wb)
      keys = task.eq_base_keys.root_key_values(true)
      add_locale_key_value_table wb, keys, name: eq_base_title(keys)
    end

    def add_unused_sheet(wb)
      keys = task.unused_keys.root_key_values(true)
      add_locale_key_value_table wb, keys, name: unused_title(keys)
    end

    private

    def add_locale_key_value_table(wb, keys, worksheet_opts = {})
      wb.add_worksheet worksheet_opts do |sheet|
        sheet.add_row [I18n.t('i18n_tasks.common.locale'), I18n.t('i18n_tasks.common.key'), I18n.t('i18n_tasks.common.value')]
        style_header sheet
        keys.each do |locale_k_v|
          sheet.add_row locale_k_v
        end
      end
    end


    def style_header(sheet)
      border_bottom = sheet.workbook.styles.add_style(border: {style: :thin, color: '000000', edges: [:bottom]})
      sheet.rows.first.style = border_bottom
    end
  end
end
