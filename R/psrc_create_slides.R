#' Create Cover Slide
#'
#' This function allows you to create a title slide for a new presentation.
#' @param p PowerPoint Template File to use to create presentation - defaults to template installed with package
#' @param p.img Location or URL for background image for cover slide - ex "images/bellevuelakewashington_3.jpg"
#' @param p.title Title for the presentation
#' @param p.mtg Meeting or presentation venue
#' @param p.date Meeting date
#' @return Cover slide for PowerPoint Presentation
#' 
#' @examples
#' 
#' psrc_pres = officer::read_pptx(system.file('extdata', 
#'                                            'psrc-trends-template.pptx', 
#'                                             package='psrcslides'))
#' 
#' psrc_pres <- add_cover_slide(p=psrc_pres,
#'                              p.title="Puget Sound Trends",
#'                              p.mtg="Growth Management Policy Board",
#'                              p.date="September 2022")
#' 
#' @export
#'
add_cover_slide <- function(p,
                            p.img=system.file('extdata', 'bellevuelakewashington.jpg', package='psrcslides'),
                            p.title="Presentation Title",
                            p.mtg="Meeting Name",
                            p.date="Meeting Date") {
  
  pres <- officer::add_slide(p, layout = "Title Slide") 
  pres <- officer::ph_with(x=pres,
                           value=officer::external_img(file.path(p.img)), 
                           location = officer::ph_location_type(type = "pic"))
  pres <- officer::ph_with(x=pres, 
                           value=p.title, 
                           location = officer::ph_location_type(type = "ctrTitle"))
  pres <- officer::ph_with(x=pres,
                           value=p.mtg, 
                           location = officer::ph_location_type(type = "subTitle"))
  pres <- officer::ph_with(x=pres, 
                           value=p.date, 
                           location = officer::ph_location_type(type = "dt"))
  return(pres)
}

#' Create Transition Slide
#'
#' This function creates a transition slide with a background picture and white centered text.
#' @param p officer presentation file to add slide
#' @param p.img Location or URL for background image - ex "images/bellevuelakewashington_3.jpg"
#' @param p.title Title for slide
#' @return Transition slide for PowerPoint Presentation
#' 
#' @examples
#' 
#' library(officer)
#' 
#' psrc_pres = read_pptx(system.file('extdata', 
#'                                   'psrc-trends-template.pptx', 
#'                                    package='psrcslides'))
#' 
#' psrc_pres <- add_transition_slide(p=psrc_pres, 
#'                                   p.title="Population Trends",
#'                                   p.img=system.file('extdata', 
#'                                                     'uw-equity.png',
#'                                                      package='psrcslides'))
#'                                   
#' @export
#'
add_transition_slide <- function(p, p.img, p.title="Slide Title") {
  
  # Create a new slide
  pres <- officer::add_slide(p, layout = "Transition Slide") 
  
  # Add background image to slide
  pres <- officer::ph_with(x=pres, 
                           value=officer::external_img(file.path(p.img)),
                           location = officer::ph_location_type(type = "pic"))
  
  # Add Slide Title
  pres <- officer::ph_with(x=pres, 
                           value=p.title, 
                           location = officer::ph_location_type(type = "title"))
  return(pres)
}

#' Create Full Text Slide
#'
#' This function creates a text slide with a large text box for bullets and a caption.
#' @param p officer presentation file to add slide
#' @param p.title Title for slide
#' @param p.caption Caption text for slide
#' @param p.text Text for bullets 
#' @return Full Text slide for PowerPoint Presentation
#' 
#' @examples
#' 
#' library(officer)
#' 
#' psrc_pres = read_pptx(system.file('extdata', 
#'                                   'psrc-trends-template.pptx', 
#'                                    package='psrcslides'))
#' 
#' psrc_pres <- add_full_text_slide(p=psrc_pres, 
#'                                  p.title="Presentation Outline",
#'                                  p.caption="Topics covered today include:",
#'                                  p.text=paste0("Population Trends\n",
#'                                                "Housing Trends\n",
#'                                                "Job Trends\n",
#'                                                "Transportation Trends"))
#'                                   
#' @export
#'
add_full_text_slide <- function(p, p.title="Slide Title", p.caption="Caption Text", p.text="Bullet Points") {
  
  # Create a new slide
  pres <- officer::add_slide(p, layout = "Title with Full Text and Caption") 
  
  # Add Slide Title
  pres <- officer::ph_with(x=pres, 
                           value=p.title, 
                           location = officer::ph_location_type(type = "title"))
  
  # Add Slide Caption above main Text
  pres <- officer::ph_with(x=pres, 
                           value=p.caption, 
                           location = officer::ph_location_label(ph_label = "Text Placeholder 3"))
  
  # Add Slide Bullet Text
  pres <- officer::ph_with(x=pres, 
                           value=p.text, 
                           location = officer::ph_location_label(ph_label = "Text Placeholder 5"))
  
  return(pres)
}

#' Create Slide with Full width picture or image and a caption
#'
#' This function creates a data slide with a full width object and a caption text.
#' @param p officer presentation object
#' @param p.title Title for slide
#' @param p.caption Text to be used for summary above the chart
#' @param p.chart Chart object to display in slide
#' 
#' @return Full Data slide with caption for PowerPoint presentation
#' 
#' @examples
#' 
#' library(dplyr)
#' library(officer)
#' library(psrctrends)
#' library(psrcplot)
#' 
#' psrc_pres = read_pptx(system.file('extdata', 
#'                                   'psrc-trends-template.pptx', 
#'                                    package='psrcslides'))
#'                                             
#' ofm.pop <- get_ofm_intercensal_population() %>% 
#'     filter(Jurisdiction=="Region") %>%
#'     mutate(Delta = round(Estimate - lag(Estimate),-2)) %>%
#'     filter(Year >=2013) %>%
#'     mutate(Year = as.character(Year))
#'     
#' pop.change <- create_bar_chart(t=ofm.pop, 
#'                                w.x="Year", 
#'                                w.y="Delta", 
#'                                f="Variable", 
#'                                w.color="GnPr",
#'                                w.title="Population Change: 2013 to 2022",
#'                                w.sub.title="Source: OFM",
#'                                est.type = "number")
#' 
#' psrc_pres <- add_full_chart_caption_slide(p = psrc_pres, 
#'                                           p.title = "Despite a Global Pandemic, 
#'                                              the region continues to grow",
#'                                           p.caption = "Enter your caption text here",
#'                                           p.chart = pop.change)
#'                                   
#' @export
#'
add_full_chart_caption_slide <- function(p, p.title="Slide Title", p.caption=" ", p.chart) {
  
  # Create a new slide
  pres <- officer::add_slide(p, layout = "Title with Full Chart and Caption") 
  
  # Add Slide Title
  pres <- officer::ph_with(x=pres,
                           value=p.title, 
                           location = officer::ph_location_type(type = "title"))
  
  # Add Slide Caption above main chart
  pres <- officer::ph_with(x=pres, 
                           value=p.caption, 
                           location = officer::ph_location_type(type = "body"))
  
  # Add Chart
  pres <- officer::ph_with(x=pres, 
                           value=p.chart, 
                           location = officer::ph_location_type(type = "pic"))
  
  return(pres)
}

#' Create Slide with Full width picture or image
#'
#' This function creates a data slide with a full width object without caption text.
#' @param p officer presentation object
#' @param p.title Title for slide
#' @param p.chart Chart object to display in slide
#' 
#' @return Full Data slide with caption for PowerPoint presentation
#' 
#' @export
#'
add_full_chart_slide <- function(p, p.title="Slide Title", p.chart) {
  
  # Create a new slide
  pres <- officer::add_slide(p, layout = "Title with Full Chart") 
  
  # Add Slide Title
  pres <- officer::ph_with(x=pres,
                           value=p.title, 
                           location = officer::ph_location_type(type = "title"))
  # Add Chart
  pres <- officer::ph_with(x=pres, 
                           value=p.chart, 
                           location = officer::ph_location_type(type = "pic"))
  
  return(pres)
}

#' Create Slide with Full Bullets on left and chart on the left
#'
#' This function creates a data slide with a bullets and a chart.
#' @param p officer presentation object
#' @param p.title Title for slide
#' @param p.bullet Text for bullet points
#' @param p.caption Text to be used for summary above the chart
#' @param p.chart Chart object to display in slide
#' 
#' @return Full Data slide with caption for PowerPoint presentation
#' 
#'                                   
#' @export
#'
add_bullet_plus_chart_slide <- function(p, p.title=" ", p.caption=" ", p.bullet=" ", p.chart) {
  
  # Create a new slide
  pres <- officer::add_slide(p, layout = "Title with Bullets, Chart and Caption") 
  
  # Add Slide Title
  pres <- officer::ph_with(x=pres,
                           value=p.title, 
                           location = officer::ph_location_type(type = "title"))
  
  # Add Slide Caption above main chart
  pres <- officer::ph_with(x=pres, 
                           value=p.caption, 
                           location = officer::ph_location_label(ph_label = "Text Placeholder 4"))
  
  # Add Bullet
  pres <- officer::ph_with(x=pres, 
                           value=p.bullet,
                           location = officer::ph_location_label(ph_label = "Text Placeholder 5"))
  
  # Add Chart
  pres <- officer::ph_with(x=pres, 
                           value=p.chart, 
                           location = officer::ph_location_type(type = "pic"))
  
  return(pres)
}