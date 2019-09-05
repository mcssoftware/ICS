class Css {
    public combine(...styles): string {
        if (Array.isArray(styles) && styles.length > 0){
            return styles.join(' ');
        }
        return '';
    }
}

export default new Css();