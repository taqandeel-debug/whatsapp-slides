import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import io

# Page Config for Mobile
st.set_page_config(page_title="WhatsApp Slide Maker", layout="wide")

st.title("📱 Daily WhatsApp Slides")
st.write("Upload your photos. I will sort them by time and put 3 on each slide.")

# 1. File Uploader (Mobile native camera/gallery)
uploaded_files = st.file_uploader(
    "Tap to add photos", 
    accept_multiple_files=True, 
    type=['png', 'jpg', 'jpeg']
)

if uploaded_files:
    # 2. Sort Logic: WhatsApp images are named by date (IMG-2023...), so sorting by name works perfectly
    uploaded_files.sort(key=lambda x: x.name)
    
    st.success(f"✅ {len(uploaded_files)} photos loaded in chronological order.")
    
    # 3. Preview Section
    st.divider()
    st.subheader("👀 Slide Preview")
    
    # Group images into chunks of 3
    chunks = [uploaded_files[i:i + 3] for i in range(0, len(uploaded_files), 3)]
    
    for i, group in enumerate(chunks):
        st.write(f"**Slide {i+1}**")
        cols = st.columns(3)
        for idx, file in enumerate(group):
            # Display image in the column
            cols[idx].image(file, use_container_width=True)
            cols[idx].caption(f"Pic {idx+1}")
    
    st.divider()

    # 4. PPT Generation Button
    if st.button("🚀 Convert to PowerPoint"):
        # Create Presentation
        prs = Presentation()
        # Set to 16:9 widescreen (ideal for phones/TVs)
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # Loop through groups and create slides
        for group in chunks:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank Layout
            
            # Layout Calculation: 3 images side-by-side
            # Slide Width: 13.33" | Image Width: ~4.2" | Spacing: ~0.2"
            positions = [Inches(0.2), Inches(4.56), Inches(8.92)]
            top = Inches(1.5) # Vertically centered
            height = Inches(4.5) # Fixed height, width scales automatically
            
            for idx, file in enumerate(group):
                # Reset file pointer to beginning so PPTX can read it
                file.seek(0)
                slide.shapes.add_picture(file, positions[idx], top, height=height)

        # Save to memory buffer
        ppt_output = io.BytesIO()
        prs.save(ppt_output)
        ppt_output.seek(0)
        
        st.download_button(
            label="📥 Download Final PPT",
            data=ppt_output,
            file_name="Daily_Updates.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
