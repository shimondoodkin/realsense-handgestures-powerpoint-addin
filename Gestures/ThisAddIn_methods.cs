using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Gestures
{
    public partial class ThisAddIn
    {


        public void gotoSlide(int addcount)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    PowerPoint.Slide slide = (PowerPoint.Slide)view.Slide;
                    if (slide.SlideIndex + addcount > 0 && addcount <= presentation.Slides.Count)
                        view.GotoSlide(slide.SlideIndex + addcount);
                }

            }
            catch (Exception)
            {

                // throw;
            }

            try
            {
                if (Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow.Active == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.SlideShowView view = Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow.View;
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    PowerPoint.Slide slide = (PowerPoint.Slide)view.Slide;
                    if (slide.SlideIndex + addcount > 0 && addcount <= presentation.Slides.Count)
                        view.GotoSlide(slide.SlideIndex + addcount);
                }

            }
            catch (Exception)
            {

                // throw;
            }

        }


        public void newslide()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    if (presentation.Slides.Count > 0)
                    {
                        PowerPoint.Slide slide = (PowerPoint.Slide)view.Slide;
                        presentation.Slides.AddSlide(slide.SlideIndex + 1, presentation.SlideMaster.CustomLayouts._Index(PowerPoint.PpSlideLayout.ppLayoutTitle.GetHashCode()));
                        presentation.Slides[slide.SlideIndex + 1].Select();
                    }
                    else
                    {
                        presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts._Index(PowerPoint.PpSlideLayout.ppLayoutTitle.GetHashCode()));
                        //presentation.Slides[0].Select();
                    }
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void deleteslide()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    PowerPoint.Slide slide = (PowerPoint.Slide)view.Slide;
                    presentation.Slides[slide.SlideIndex].Delete();
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void redo()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Redo");
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void undo()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void settext(string title)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    PowerPoint.Slide slide = (PowerPoint.Slide)view.Slide;
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void startshow()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Active == Office.MsoTriState.msoTrue)
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("SlideShowFromCurrent");
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }

        public void endshow()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow.Active == Office.MsoTriState.msoTrue)
                {   
                    PowerPoint.SlideShowView view = Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow.View;
                    view.Exit();
                    //Globals.ThisAddIn.Application.ActivePresentation.Windows(0).Activate();
                }
            }
            catch (Exception)
            {

                // throw;
            }
        }
    }
}
