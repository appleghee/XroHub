document.addEventListener('DOMContentLoaded', () => {
    console.log('XroHub loaded');
    
    // Toggle sidebar functionality could be added here
    const menuBtn = document.querySelector('header button');
    const sidebar = document.querySelector('aside');
    const main = document.querySelector('main');

    if (menuBtn && sidebar && main) {
        menuBtn.addEventListener('click', () => {
            sidebar.classList.toggle('hidden');
            main.classList.toggle('lg:ml-60');
        });
    }
});
