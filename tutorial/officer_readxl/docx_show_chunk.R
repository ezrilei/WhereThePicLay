officer_doc <- officer::read_docx(path = "./Data/【模板】慰问信.docx")
officer_doc %>% cursor_begin() %>% 
  cursor_forward %>% 
  cursor_forward %>% 
  #cursor_forward %>% 
  #cursor_forward %>% 
  #cursor_forward %>%
  docx_show_chunk()

officer_doc %>% cursor_begin() %>% 
  cursor_forward %>% 
  cursor_forward %>% 
  cursor_forward %>% 
  cursor_forward %>% 
  #cursor_forward %>%
  docx_show_chunk()

#————————————————————————————————
# def variable
namePerson <- tt.info$姓名[1]
nameBank <- tt.info$开户行名称[1]
account <- tt.info$`开户行账号（卡号）`[1] %>% as.character()

# 替换
doc_replaced <- officer_doc %>%
  body_replace_all_text(old_value = "namePerson", 
                        new_value = namePerson, 
                        only_at_cursor = FALSE
  ) %>%
  body_replace_all_text(old_value = "nameBank", 
                        new_value = nameBank, 
                        only_at_cursor = FALSE
  ) %>% 
  body_replace_all_text(old_value = "account", 
                        new_value = account, 
                        only_at_cursor = FALSE
  )