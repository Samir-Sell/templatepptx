
class TemplatePptxOptions:
    '''
    Description: Manage the options around templatepptx
    '''

    def __init__(self):
        self._strict = False

    @property
    def strict_mode(self) -> bool:
        '''
        Description: Return the value of strict mode. Strict mode
        disables running through warnings and will cause the application to be stopped
        and raise an error.
        '''
        return self._strict
    
    @strict_mode.setter
    def strict_mode(self, strict_enabled: bool) -> None:
        '''
        Description: Set the strict mode of the TemplatePptx process. Strict mode
        disables running through warnings and will cause the application to be stopped
        and raise an error.

        @input strict_enable: A boolean indicating if strict mode is enabled or not
        '''
        self._strict = strict_enabled