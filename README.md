# Proof-of-Concept Excel Add-In for Peek Statistics

## Overview
This add-in allows Peek stats to be embeded in a spreadsheet and refreshed on demand without using macros.

Cells are 'tagged' for stats using range names that start with 'PEEK'. The names may encode queries or reference canned queries.

Currently stats are simply counts of encounters or orders with selection criteria.

## Encoded Queries

Here's an example

    PEEK.kenyaschools.test.Encounters.type.vision_screening._observations__gender.female
  
The rules are simple. Separate fields with a "."

* PEEK
* Project (kenyaschools, botswanaschools etc.)
* Environment (test, preprod, prod)
* Collection (Encounters, Orders)
* Predicates expressed as name.value pairs

For a property name, '__' is replaced by a '.'

For a For property value, a leading '__' is interpreted as a "not equal"


## Canned Queries

Defined the same way in code but referenced with a short code.


      'PEEK.KSF': Kenya screenings female     'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.female',
      'PEEK.KSM': Kenya screenings male       'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.male',
      'PEEK.KPF': Kenya +ve screenings female 'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.female._observations__healthy_eyes.false',
      'PEEK.KPM': Kenya +ve screenings male   'PEEK.kenyaschools.prod.Encounters.type.vision_screening._observations__gender.male._observations__healthy_eyes.false',
      'PEEK.KTF': Kenya triage female         'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female',
      'PEEK.KTM': Kenya triage male           'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male',
      'PEEK.KRF': Kenya refraction female     'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
      'PEEK.KRM': Kenya refraction male       'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
      'PEEK.KGF': Kenya glasses female        'PEEK.kenyaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.female',
      'PEEK.KGM': Kenya glasses male          'PEEK.kenyaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.male',
      'PEEK.KOF': Kenya non-refractive female 'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',
      'PEEK.KOM': Kenya non-refractive male   'PEEK.kenyaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',
  
      'PEEK.BSF': Botsw screenings female     'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.female',
      'PEEK.BSM': Botsw screenings male       'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.male',
      'PEEK.BPF': Botsw +ve screenings female 'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.female._observations__healthy_eyes.false',
      'PEEK.BPM': Botsw +ve screenings male   'PEEK.botswanaschools.prod.Encounters.type.vision_screening._observations__gender.male._observations__healthy_eyes.false',
      'PEEK.BTF': Botsw triage female         'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female',
      'PEEK.BTM': Botsw triage male           'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male',
      'PEEK.BRF': Botsw refraction female     'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
      'PEEK.BRM': Botsw refraction male       'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__triage_outcome_refraction.__triage_outcome_refraction_none',
      'PEEK.BGF': Botsw glasses female        'PEEK.botswanaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.female',
      'PEEK.BGM': Botsw glasses male          'PEEK.botswanaschools.prod.Orders.type.spectacles_order.status.order_status_dispensed.gender.male',
      'PEEK.BOF': Botsw non-refractive female 'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.female._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none',
      'PEEK.BOM': Botsw non-refractive male   'PEEK.botswanaschools.prod.Encounters.type.vision_triage.status.finished.gender.male._observations__healthy_eyes.false._observations__triage_outcome_refraction.triage_outcome_refraction_none'

## Constraints

Names were chosen as maintainable cell metatdata. However they are limited in composition (hence the use of __) and length (a major reason for for canned queries).

Using comments sounds like an attractive altenative but comments are not (yet?) accessible through the JavaScript API.

## Next Steps

* Make canned queries cross-project and environment
* Extend beyond scalar counts to aggregates (group API) an
* Extend beyond aggregates to raw data



