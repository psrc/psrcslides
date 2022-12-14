% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/psrc_create_slides.R
\name{add_full_text_slide}
\alias{add_full_text_slide}
\title{Create Full Text Slide}
\usage{
add_full_text_slide(
  p,
  p.title = "Slide Title",
  p.caption = "Caption Text",
  p.text = "Bullet Points"
)
}
\arguments{
\item{p}{officer presentation file to add slide}

\item{p.title}{Title for slide}

\item{p.caption}{Caption text for slide}

\item{p.text}{Text for bullets}
}
\value{
Full Text slide for PowerPoint Presentation
}
\description{
This function creates a text slide with a large text box for bullets and a caption.
}
\examples{

library(officer)

psrc_pres = read_pptx(system.file('extdata', 
                                  'psrc-trends-template.pptx', 
                                   package='psrcslides'))

psrc_pres <- add_full_text_slide(p=psrc_pres, 
                                 p.title="Presentation Outline",
                                 p.caption="Topics covered today include:",
                                 p.text=paste0("Population Trends\n",
                                               "Housing Trends\n",
                                               "Job Trends\n",
                                               "Transportation Trends"))
                                  
}
