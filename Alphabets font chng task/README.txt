Process to create PPT through Python script in Task 1-

STEP-1: Install the Python library of pptx by the system's current Python version by following the command: ( pip install python-pptx)

STEP-2: Install the sample_font_file.ttf in your system to get access to it in our presentation.

STEP-3: -Importing collections and collections.abc to access the abstract base classes from it.
        -From pptx library Import Presentation
        -From pptx.util import Inches, Pt

STEP-4: Declaring and Defining a function add_font_to_presentation taking our font name and presentation(here ppt) as its parameter
      
      -The function iterates through every slide in our presentation, then every shape in the presentation then checks if the shape has 
       text_frame in it, if the shape has text_frame in it then will iterate through every paragraph in it and then iterates through the  
       run part of the paragraph assigning run part the font size, style, and other required attributes.
        
       Hence, assigning our content with the required font styling.

STEP-5: Our code starts here-
     
       -giving the reference of Presentation() to any object(here ppt)
       -giving reference to slide_layout1(any object) of the slide layout we want our slide to have by ppt.slide_Layouts[6]
       -adding slides to a presentation by ppt.slides.add_slide(slide_layout1)
       -adding text box to slides by shapes.add_textbox(left,top,width,height)
       -adding text_frame to the text box to access it later in the code
       -opening given sample_slide2_input.txt type file and reading its content and passing it to the contents variable then adding   
        content to our text box on the slide
       -wrapping the text of the text box so it will not go outside the slide by setting word_wrap=True
       -Giving the value to parameters of our function defined earlier
       -Finally save our ppt by ppt.save('PPT.pptx')