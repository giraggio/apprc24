library(arrow)
localidades <- read_parquet("Censo/2024/Cartografia_censo2024_Pais/Cartografia_censo2024_Pais_Localidades.parquet")
library(dplyr)
library(sf)

localidades_comunas <- localidades |>
  filter(COMUNA == "SIERRA GORDA")

library(sf)

localidades_comuna_sf <- localidades_comunas |>
  st_as_sf(crs = 4326)
localidades_comuna_sf


library(ggplot2)

localidades_comuna_sf |>
  ggplot() +
  aes(fill = n_fuente_agua_publica) +
  geom_sf(color = "white", linewidth = 0.02) +
  scale_fill_fermenter(palette = 12) +
  theme_minimal(base_size = 10) +
  theme(axis.text.x = element_text(angle = 90, vjust = .5)) +
  guides(fill = guide_legend(title = "Población",
                             position = "top")) +
  labs(title = "Localidades por Población",
       subtitle = "Comuna de San Pedro",
       caption = "Fuente: Censo 2024, INE")


# mapa interactivo

library(mapgl)

localidades_comuna <- localidades |>
  filter(COMUNA == "SAN PEDRO") |>
  select(REGION, n_per, SHAPE)

localidades_comuna_sf <- localidades_comuna |>
  st_as_sf(crs = 4326)

localidades_comuna_sf |>
  maplibre_view(column = "n_fuente_agua_publica")

