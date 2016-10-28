(function () {
    'use strict';
    const CANNED_QUERIES = {
        'PEEK.KSF': 'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.female',
        'PEEK.KSM': 'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.male',
        'PEEK.KPF': 'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.female._observations__healthy_eyes.false',
        'PEEK.KPM': 'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.male._observations__healthy_eyes.false',
        'PEEK.KTF': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female',
        'PEEK.KTM': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male',
        'PEEK.KRF': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
        'PEEK.KRM': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
        'PEEK.KGF': 'PEEK.kenyaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.female',
        'PEEK.KGM': 'PEEK.kenyaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.male',
        'PEEK.KOF': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',
        'PEEK.KOM': 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',

        'PEEK.BSF': 'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.female',
        'PEEK.BSM': 'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.male',
        'PEEK.BPF': 'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.female._observations__healthy_eyes.false',
        'PEEK.BPM': 'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.male._observations__healthy_eyes.false',
        'PEEK.BTF': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female',
        'PEEK.BTM': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male',
        'PEEK.BRF': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
        'PEEK.BRM': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
        'PEEK.BGF': 'PEEK.botswanaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.female',
        'PEEK.BGM': 'PEEK.botswanaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.male',
        'PEEK.BOF': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',
        'PEEK.BOM': 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none'
    };

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#peek-refresh-stats').click(refreshStats);
            app.showNotification('INFO', 'Welcome to the Peek Stats Add-in');
        });
    };

    function refreshStats() {
        return Excel.run(ctx => {

            app.showNotification('INFO', 'Statistics are being refreshed...');
            // Load all namedItems.
            return ctx.sync(ctx.workbook.names.load(['items']))

            // Load the name for each item. (Note [...x] converts array-like-object to an array.)
                .then(namedItems =>[...namedItems.items].map(item => item.load('name')))
                .then(ctx.sync)

                // Update the values for PEEK ranges only
                .then(namedItems =>
                    Promise.all(namedItems
                        .filter(item => /^PEEK/.test(item.name))
                        .map(item => get(item.name)
                            .then(res => item.getRange().values = [[res.count]])
                            .catch(err => console.log('Error', err)))
                    ))
                .then(() => app.showNotification('INFO', 'Statistics have been refreshed...'))
                .catch(error => {
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                    console.log("Error: " + error);

                });
        })
    }

    /**
     *  Maps and encoded name to a URL. E.g
     *
     *  PEEK.kenyaschools.test.Encounters.type.vision_screening._observations__gender.female =>
     *  https://kenyaschools.test.peek.vision/api/Encounters/count?query=%7B%22type%22%3A%22vision_screening%22%2C%22_observations.gender%22%3A%22female%22%7D
     *
     *  For property name, a  '__' is replaced by a '.'
     *  For property value, a leading '__' is interpreted as a "not equal"
     *
     * @param name encoded as PEEK.<project>.<env>.<collection>[.<property>.<value>...]
     * @returns {*} URL
     */

    function mapNameToUrl(name) {
        let parts = (CANNED_QUERIES[name] || name).split('.'), where = {};
        parts.shift();
        let project = parts.shift();
        let env = parts.shift();
        env = env.toLowerCase() === 'prod' ? '' : `.${env}`;
        let collection = parts.shift().toLowerCase().replace(/^./, l => l.toUpperCase());
        while (parts.length > 1) {
            let property = parts.shift().replace('__', '.'), value = parts.shift();
            where[property] = /^__/.test(value) ? {neq: value.substr(2)} : value;
        }
        console.log(name, '\n', where );
        return `https://${project}${env}.peek.vision/api/${collection}/count?where=${encodeURIComponent(JSON.stringify(where))}`;
    }

    function get(name) {
        let url;
        return Promise.resolve()
            .then(()=>mapNameToUrl(name))
            .then(res => url=res)
            .then(url => $.ajax({url}))
            // .then(res => {
            //     console.log(name, url, res);
            //     return res;
            // });
    }
})();

