#Script author - Sergey Rozum, https://www.uxrozum.com/
#If you are seeing R for the first time, this material will be very useful - https://thecode.media/rrrrr/ , 
  # https://www.youtube.com/playlist?list=PLRDJTUEth2lErO5xbwoTyl9-X3H7lokZI

#Step 0: Install libraries (run only once after installing R)
install.packages(c('openxlsx','igraph', 'factoextra', "ggwordcloud", 'rstudioapi'))

# Step 1: Load libraries
library(igraph)
library(openxlsx)
library(factoextra)
library(ggwordcloud)
library(rstudioapi)

#Step 2: Select a folder
setwd(
  rstudioapi::selectDirectory()) #This should be a separate folder on your computer
#with a Latin name
#The folder should contain a data file - Card.xlsx

#Step 2.5: Load data
Raw <- read.xlsx('Card.xlsx') #The file must have at least three columns
#Card - card names
#Group_id - group ID (unique for EACH group)
#Group_name - group names provided by respondents

#Step 3: Create an adjacency table
Adj <- crossprod(table(Raw$Group_id, Raw$Card))
diag(Adj) = 0

#Step 4: Save the adjacency table
write.xlsx(as.data.frame(Adj), 'Adjacency.xlsx',
           overwrite = T, col.names = T, row.names=T)

#Step 5: Clustering

#Edge betweenness algorithm - https://en.wikipedia.org/wiki/Girvan%E2%80%93Newman_algorithm
Net <-  graph_from_adjacency_matrix(Adj, mode='undirected')
#Build network structure based on the adjacency table
Clust <- as.dendrogram(cluster_edge_betweenness(Net)) #Build clusters
fviz_dend(Clust, k=6 #Number of groups to obtain
          ,horiz=T) #Display the plot

#k-means algorithm - https://en.wikipedia.org/wiki/K-means_clustering
fviz_nbclust(Adj, kmeans) #Calculate the optimal number of clusters
km <- kmeans(Adj , centers = 6 #Build clusters
             , nstart = 25)
fviz_cluster(km, data= Adj) #Display the plot

#Hierarchical clustering algorithm (Ward) - https://en.wikipedia.org/wiki/Ward%27s_method

plot(hclust(dist(Adj), method='ward.D')) #Simultaneously form clusters and display the plot

#Step 6: Build a network graph
plot.new()
plot.igraph(
  graph_from_adjacency_matrix(ifelse(Adj-10 < 0, 0, Adj-10)
                              #This is the number of connections to filter out, 
                              #the higher the number, the stronger the connections
                              #if the graph looks cluttered, increase this number
                              #focus on the number of respondents
                              
                              , mode='undirected'), 
  vertex.label.color= "black", vertex.color= "gray", 
  vertex.size= 20, vertex.frame.color='gray',asp = 0.7,
  layout = layout.kamada.kawai,
  vertex.label.cex = 0.9,
  width=1, height=1)

#Step 7: Word clouds for group names

cardgroup <- c('Cockroach', 'Spider', 'Dragonfly') #Write card names here
gnames <- as.data.frame(table(Raw$Group_name[Raw$Card %in%  cardgroup])) #Extract card names
ggplot(gnames, aes(label = Var1, size= Freq, color = Freq)) +
  geom_text_wordcloud() +
  scale_size_area(max_size = 50 #Maximum word size; if all words do not fit, reduce this value
  ) +
  theme_minimal() 

#Quick option: Select everything between start and end, and press ctrl+enter
# Important - The file with the cards must be named cards.xlsx

# If not installed, install 
install.packages(c('openxlsx','igraph', 'factoextra', "ggwordcloud", 'rstudioapi'))

#Start
library(igraph)
library(openxlsx)
library(factoextra)
library(ggwordcloud)
library(rstudioapi)

setwd(rstudioapi::selectDirectory())
Raw <- read.xlsx('Card.xlsx')
Adj <- crossprod(table(Raw$Group_id, Raw$Card))
diag(Adj) = 0
write.xlsx(as.data.frame(Adj), 'Adjacency.xlsx',
           overwrite = T, col.names = T, row.names=T)
Net <-  graph_from_adjacency_matrix(Adj, mode='undirected')
Clust <- as.dendrogram(cluster_edge_betweenness(Net))
png(filename="Dendro_%02d.png",
    width = 1000, height = 1000, units = "px")
for(i in 1:length(unique(Raw$Card))) {
  myPlot <- fviz_dend(Clust, k=i,horiz=T,
                      main=paste("Dendrogram of card distribution by groups \n",
                                 "Number of groups:", i))
  print(myPlot)
}
dev.off()
for(i in 0:19) {
  Adj2 <- Adj-max(Adj)*i/20
  png(filename=paste('Net graph ', i ,'.png'),
      width = 1000, height = 1000, units = "px")
  myPlot <- plot.igraph(
    graph_from_adjacency_matrix(ifelse(Adj2 < 0, 0, Adj2), mode='undirected'),
    vertex.label.color= "black", vertex.color= "gray", 
    vertex.size= 20, vertex.frame.color='gray',asp = 0.7,
    layout = layout.kamada.kawai,
    vertex.label.cex = 0.9,
    width=1, height=1)
  print(myPlot+title(paste('Network graph of card distribution by groups \n',
                           100-(i*5), '% connections'),cex.main=2))
  dev.off()
}

for(i in 1:length(table(Raw$Card))) {
  png(filename=paste('Wordcloud ', i ,'.png'),
      width = 1000, height = 1000, units = "px")
  gnames <- as.data.frame(table(Raw$Group_name[Raw$Card == as.data.frame(table(Raw$Card))$Var1[i]]))
  myPlot <- ggplot(gnames, aes(label = Var1, size= Freq, color = Freq)) +
    geom_text_wordcloud() +
    scale_size_area(max_size = 20) +
    ggtitle(paste('Word cloud for card: ', as.data.frame(table(Raw$Card))$Var1[i])) +
    theme_minimal() + 
    theme(plot.title = element_text(hjust = 0.5, size=40, margin = unit(c(20,0,0,0), 'pt')))
  print(myPlot)
  dev.off()
}
#End
