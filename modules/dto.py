''' for data transfer objects '''
from dataclasses import dataclass
@dataclass(frozen=True)
class UseCaseDataDTO:
    ''' DTO object for vars exchange '''
    def __init__(self, company, job_type, position, job_portal, replacements):
        object.__setattr__(self, 'company', company)
        object.__setattr__(self, 'job_type', job_type)
        object.__setattr__(self, 'position', position)
        object.__setattr__(self, 'job_portal', job_portal)
        object.__setattr__(self, 'replacements', replacements)