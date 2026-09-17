const projectPhotos = [
    {
        project: 'McNair Campus',
        image: 'media/McNair/Exterior.jpg',
        alt: 'Exterior view of the McNair campus with a covered entry and landscaped approach',
        story: 'A clear arrival experience helps patients, families, and staff feel oriented from the first step.'
    },
    {
        project: 'McNair Campus',
        image: 'media/McNair/Corridor.jpg',
        alt: 'Bright clinical corridor with repeated doorways and directional signage',
        story: 'Operational thinking turns circulation into a calmer, more intuitive part of care delivery.'
    },
    {
        project: 'McNair Campus',
        image: 'media/McNair/Stations.jpg',
        alt: 'Clinical workstations arranged along a bright patient care area',
        story: 'The best environments support the teams who use them, keeping work visible, efficient, and connected.'
    },
    {
        project: 'McNair Campus',
        image: 'media/McNair/Lobby.jpg',
        alt: 'Spacious healthcare lobby with seating, artwork, and a reception desk',
        story: 'Details in shared spaces build confidence and make complex facilities feel human.'
    },
    {
        project: 'McNair Campus',
        image: 'media/McNair/OutpatientImagingDesk.jpg',
        alt: 'Outpatient imaging reception desk with a clean, welcoming layout',
        story: 'Thoughtful planning balances the needs of patients, clinical teams, and the systems behind the scenes.'
    },
    {
        project: 'McNair Campus',
        image: 'media/McNair/AboveCeiling.jpg',
        alt: 'Building systems visible above a ceiling grid during construction',
        story: 'Long-term performance starts above the ceiling, where coordination protects reliability and maintainability.'
    },
    {
        project: 'Fort Lauderdale Airport',
        image: 'media/FLL Airport/09-28-2005_2.jpg',
        alt: 'Airport terminal interior with structural columns and passenger circulation space',
        story: 'Experience across demanding public environments sharpens our ability to coordinate complex work.'
    },
    {
        project: 'Fort Lauderdale Airport',
        image: 'media/FLL Airport/DSCF0650.JPG',
        alt: 'Airport terminal concourse with an open passenger waiting area',
        story: 'Every successful project connects the finished place to the larger systems and communities around it.'
    },
    {
        project: 'Fort Lauderdale Airport',
        image: 'media/FLL Airport/RacPhotos 051.jpg',
        alt: 'Airport facility exterior viewed from the surrounding roadway',
        story: 'High-traffic spaces depend on disciplined phasing, clear communication, and resilient infrastructure.'
    },
    {
        project: 'Office Depot Global Headquarters',
        image: 'media/OfficeDepotGlobalHeadquarters/OD_EXT_D_IMG_9267.jpg',
        alt: 'Exterior view of the Office Depot Global Headquarters with a landscaped entrance',
        story: 'A strong workplace begins with an arrival that reflects the organization and welcomes its people.'
    },
    {
        project: 'Office Depot Global Headquarters',
        image: 'media/OfficeDepotGlobalHeadquarters/OD_INT_A_IMG_7529.jpg',
        alt: 'Interior office space at the Office Depot Global Headquarters with open work areas',
        story: 'Flexible, well-coordinated interiors help teams work comfortably as needs change over time.'
    },
    {
        project: 'Office Depot Global Headquarters',
        image: 'media/OfficeDepotGlobalHeadquarters/OD_INT_C_IMG_8675.jpg',
        alt: 'Office Depot Global Headquarters interior showing a finished collaborative space',
        story: 'The details of a finished environment turn a complex build into a place people can use with confidence.'
    }
];

const gallery = document.querySelector('[data-gallery]');

projectPhotos.forEach(({ project, image, alt, story }, index) => {
    const item = document.createElement('article');
    item.className = 'gallery-item';
    item.innerHTML = `
        <button class="gallery-trigger" type="button" aria-expanded="false" aria-controls="gallery-story-${index}">
            <img src="${image}" alt="${alt}" loading="lazy" decoding="async">
            <span class="gallery-story" id="gallery-story-${index}">
                <span class="gallery-project">${project}</span>
                <span class="gallery-value">${story}</span>
                <span class="gallery-action" aria-hidden="true">View project value</span>
            </span>
        </button>
    `;
    gallery.appendChild(item);
});

gallery.addEventListener('click', (event) => {
    const trigger = event.target.closest('.gallery-trigger');
    if (!trigger) return;

    const expanded = trigger.getAttribute('aria-expanded') === 'true';
    trigger.setAttribute('aria-expanded', String(!expanded));
    trigger.closest('.gallery-item').classList.toggle('is-active', !expanded);
});

window.addEventListener('scroll', () => {
    gallery.querySelectorAll('.gallery-item.is-active').forEach((item) => {
        item.classList.remove('is-active');
        item.querySelector('.gallery-trigger').setAttribute('aria-expanded', 'false');
    });
}, { passive: true });