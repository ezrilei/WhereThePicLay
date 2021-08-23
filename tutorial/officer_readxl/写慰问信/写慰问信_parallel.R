#### set working directory ####
# set the folder where the RScript locates as working dir (you need Rstudio to do this)
setwd(
  dirname(rstudioapi::getSourceEditorContext()$path)
)

#### load Packages ####
library(readxl)
library(officer)
library(stringr)
library(parallel) # <----------

#### create output dir ####
current.time <- Sys.time() # 先把时间记下来 免得程序跑得过久 前后命名不一致
name.dir <- paste0("output_",
                   format(current.time, "%m%d_%H%M")
)
dir.create(name.dir)

path.tmp.dir <- paste(".", name.dir, "tmp/", sep = "/")
dir.create(path.tmp.dir)


### load tt.info <-----------------------
path.tt.info <- "./Data/【模拟】生活困难党员摸底统计.xlsx"

tt.info <- readxl::read_excel(path = path.tt.info,
                              sheet = 1, skip = 0
                              )


#### def  word documents Function ####

Fun.writeDoc <- function(i){
  
  # 读取 模板.docx <-----------------------
  officer_doc <- officer::read_docx(path = "./Data/【模板】慰问信.docx")
  
  # def variable
  namePerson <- tt.info$姓名[i]
  nameBank <- tt.info$开户行名称[i]
  account <- tt.info$`开户行账号（卡号）`[i] %>% as.character()
  
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
  #"./output_%m%d_%H%M/tmp/【宋江】慰问信.docx"
  print(doc_replaced, 
        target = paste0(path.tmp.dir, i,"【", namePerson, "】", "慰问信.docx")
  )
  
  # progress msg
  print(
    paste0(i,"【", namePerson, "】", "慰问信.docx","  created ✓")
    )
} # End of def Fun.writeDoc

#### def Fun.loadPackages ####
Fun.loadPackages <- function(packages) {
  for (p in packages) {
    library(p, character.only=TRUE)
  }
  NULL
} 


#### run parLapply ####
system.time({
  
  parallel.cl <- makeCluster(12) # 初始化12核心集群
  clusterExport(parallel.cl, "tt.info") #将该data.frame导入所有core
  clusterExport(parallel.cl, "name.dir")
  clusterExport(parallel.cl, "path.tmp.dir")
  clusterCall(parallel.cl, Fun.loadPackages, c("readxl", "officer", "stringr"))
  
  i = 1:nrow(tt.info)
  parLapply(cl = parallel.cl, X = i, fun = Fun.writeDoc) # lapply的并行版本
  
  stopCluster(parallel.cl) # 关闭集群
  
})
####_____________________####

#### merge letters into a single word doc ####
# get ordered
doc.path <- paste0(path.tmp.dir, 
                   list.files(path = path.tmp.dir)
)
idx <- doc.path %>% str_extract_all(pattern = "(?<=tmp/).+?(?=【)") %>% unlist() %>% as.numeric() %>% order() 
doc.path <- doc.path[idx]

# merge doc
doc_merge <- officer::read_docx(path = doc.path[1])

for(i in 2:length(doc.path)){
  doc_merge <- doc_merge %>%
    body_add_break() %>% # add a page break into an rdocx object
    body_add_docx(src = doc.path[i]) # add content of a docx into an rdocx object
  
  print(
    paste0(i, "/", length(doc.path), "  merged ✓")
  )
}


print(doc_merge, 
      target = paste0(name.dir,"/", "【", nrow(tt.info), "】", "慰问信.docx")
)