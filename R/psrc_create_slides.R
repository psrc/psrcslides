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
#' psrc_pres <- add_cover_slide(p.title="Puget Sound Trends",
#'                              p.mtg="Growth Management Policy Board",
#'                              p.date="September 2022")
#' 
#' @export
#'
add_cover_slide <- function(p=officer::read_pptx(system.file('extdata', 'psrc-trends-template.pptx', package='psrcslides')),
                            p.img=system.file('extdata', 'bellevuelakewashington.jpg', package='psrcslides'),
                            p.title="Presentation Title",
                            p.mtg="Meeting Name",
                            p.date="Meeting Date") {
  
  psrcplot::install_psrc_fonts()
  
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

