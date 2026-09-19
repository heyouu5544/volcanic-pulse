# Process

## How I used AI

I used ChatGPT and Codex throughout this project for brainstorming, understanding the assignment requirements, reading the template code, debugging Python, checking the dataset, and developing the visualisation.

I did not ask AI to generate the entire project at once. I worked in stages: first choosing the phenomenon, then finding and verifying a real dataset, checking its fields, creating a first working plot, and gradually changing the visual design.

AI was especially useful because I am still learning programming. It helped me understand what the Python code was doing and how the data could be translated into visual elements.

## Something I kept

One idea I kept was using VEI to control the visual strength of each eruption.

Each circle represents one volcanic eruption. Lower VEI values are shown with lighter and more muted colours, while stronger eruptions use warmer colours and larger circles. Eruptions without a recorded VEI are kept in the visualisation as hollow circles instead of being deleted.

I kept this approach because VEI is part of the original dataset and directly relates to volcanic explosiveness. This means the visual differences are connected to real information rather than being purely decorative.

## Something I rejected or changed

One of the first AI-assisted approaches was to download the Smithsonian Global Volcanism Program data automatically through its WFS service.

In practice, the request did not return a usable CSV file. I did not continue by guessing the column names or pretending that the data had downloaded successfully.

I then tried the official Smithsonian Excel download URL, but scripted downloading returned an HTTP 403 Forbidden error. Because of this, I rejected the automatic-download approach and manually downloaded the official Confirmed Holocene Eruptions spreadsheet from the Smithsonian GVP website.

The original Excel file was then placed in the `data/` folder without modification.

This was important because it showed me that code or AI suggestions can appear reasonable but still fail when tested with a real external data source.

## Changes to the visualisation

My first working visualisation used data from 1960–2025. It contained 2,229 eruptions, but the result was visually crowded.

There was also an early mistake in the y-axis mapping. The year had accidentally been included in the vertical position, producing a diagonal band instead of showing the order of eruptions within each year. I corrected this so that the y-axis now starts from 1 for each year.

I experimented with radial pulse symbols to make the image feel more connected to volcanic eruptions. However, when every eruption used a complex burst symbol, the image became too visually noisy.

I therefore simplified the design and changed the time range to 2000–2025. The final version contains 941 eruptions, including 27 records with unknown VEI.

I also changed the aspect ratio and replaced stretched elliptical symbols with true circular markers. This made the image easier to read while keeping the same data logic.

## What I learned

The most important part of this project was not only producing a final picture, but checking whether every step was actually supported by the data.

AI helped me write and debug code, but several parts still required verification. The failed data downloads, incorrect y-axis, and overly complex visual symbols all looked reasonable at first but had to be tested and changed.

This project made me more aware that using AI for programming still requires understanding the data, checking the output, and deciding whether the result actually communicates what I intended.
