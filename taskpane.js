Office.onReady(() => {
    // Office is ready
});

async function showMessage() {
    try {
        await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            const slide = slides.getSelected();
            const shapes = slide.shapes;
            const selectedShapes = shapes.getSelected();

            selectedShapes.load("items");
            await context.sync();

            if (selectedShapes.items.length === 0) {
                return alert("Please select at least one shape.");
            }

            // Dummy guideline positions (you can replace this with real logic if using guides API in future)
            const guidelineTop = 50;
            const guidelineLeft = 50;
            const guidelineBottom = 400;
            const guidelineRight = 600;

            selectedShapes.items.forEach(shape => {
                shape.load(["left", "top", "height", "width"]);
            });

            await context.sync();

            // Align each shape
            selectedShapes.items.forEach(shape => {
                // Align to top
                shape.top = guidelineTop;

                // Align to bottom
                // shape.top = guidelineBottom - shape.height;

                // Align to left
                // shape.left = guidelineLeft;

                // Align to right
                // shape.left = guidelineRight - shape.width;
            });

            await context.sync();
            alert("Shapes aligned!");
        });
    } catch (error) {
        console.error(error);
        alert("Error: " + error.message);
    }
}
